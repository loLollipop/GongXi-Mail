import assert from 'node:assert/strict';
import test from 'node:test';
import Fastify from 'fastify';
import pino from 'pino';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
process.env.JWT_SECRET = 'test-jwt-secret-for-email-password';

const { encrypt } = await import('../../lib/crypto.js');
const { AppError } = await import('../../plugins/error.js');
const { emailService, getEmailPasswordWithDb } = await import('./email.service.js');
const { default: emailRoutes } = await import('./email.routes.js');

const CLEAR_PASSWORD = 'plain-password-marker';
const CIPHERTEXT_MARKER = 'ciphertext-password-marker';

void test('getPassword validates the ID, selects only password, handles empty values, and sanitizes decryption errors', async () => {
    let row: { password: string | null } | null = { password: encrypt(CLEAR_PASSWORD) };
    const calls: unknown[] = [];
    const db = {
        emailAccount: {
            findUnique: async (input: unknown) => {
                calls.push(input);
                return row;
            },
        },
    } as unknown as Parameters<typeof getEmailPasswordWithDb>[0];

    await assert.rejects(getEmailPasswordWithDb(db, 0), { code: 'INVALID_ID', statusCode: 400 });
    await assert.rejects(getEmailPasswordWithDb(db, 1.5), { code: 'INVALID_ID', statusCode: 400 });
    assert.equal(calls.length, 0);

    assert.deepEqual(await getEmailPasswordWithDb(db, 7), { password: CLEAR_PASSWORD });
    assert.deepEqual(calls[0], {
        where: { id: 7 },
        select: { password: true },
    });

    row = { password: null };
    assert.deepEqual(await getEmailPasswordWithDb(db, 8), { password: null });

    row = null;
    await assert.rejects(getEmailPasswordWithDb(db, 9), { code: 'NOT_FOUND', statusCode: 404 });

    row = { password: CIPHERTEXT_MARKER };
    await assert.rejects(getEmailPasswordWithDb(db, 10), (error: unknown) => {
        assert.ok(error instanceof AppError);
        assert.equal(error.code, 'PASSWORD_UNAVAILABLE');
        assert.equal(error.statusCode, 500);
        assert.doesNotMatch(error.message, new RegExp(`${CLEAR_PASSWORD}|${CIPHERTEXT_MARKER}`));
        return true;
    });
});

void test('password and legacy secret routes enforce SUPER_ADMIN and password audit logs contain identifiers only', async (t) => {
    const lines: string[] = [];
    const app = Fastify({
        disableRequestLogging: true,
        loggerInstance: pino({}, { write: (line: string) => lines.push(line) }),
    });
    app.decorate('authenticateJwt', async (request) => {
        if (request.headers.authorization !== 'Bearer valid') {
            throw new AppError('UNAUTHORIZED', 'Authentication required', 401);
        }
        request.user = {
            id: 42,
            username: 'auditor',
            role: request.headers['x-test-role'] === 'super' ? 'SUPER_ADMIN' : 'ADMIN',
        };
    });
    app.decorate('requireSuperAdmin', async (request) => {
        if (request.user?.role !== 'SUPER_ADMIN') {
            throw new AppError('FORBIDDEN', 'Super admin access required', 403);
        }
    });

    const passwordMock = t.mock.method(emailService, 'getPassword', async () => ({ password: CLEAR_PASSWORD }));
    const detailMock = t.mock.method(emailService, 'getById', async (_id: number, includeSecrets = false) => ({
        id: 7,
        email: 'target@example.com',
        clientId: 'client-id',
        ...(includeSecrets ? { password: CLEAR_PASSWORD, refreshToken: CIPHERTEXT_MARKER } : {}),
    }) as never);
    await app.register(emailRoutes, { prefix: '/admin/emails' });

    const passwordUrl = '/admin/emails/7/password';
    assert.equal((await app.inject({ method: 'GET', url: passwordUrl })).statusCode, 401);
    assert.equal((await app.inject({ method: 'GET', url: passwordUrl, headers: { authorization: 'Bearer valid' } })).statusCode, 403);
    assert.equal(passwordMock.mock.callCount(), 0);

    const superHeaders = { authorization: 'Bearer valid', 'x-test-role': 'super' };
    const passwordResponse = await app.inject({ method: 'GET', url: passwordUrl, headers: superHeaders });
    assert.equal(passwordResponse.statusCode, 200);
    assert.equal(passwordResponse.headers['cache-control'], 'no-store');
    assert.equal(passwordResponse.headers.pragma, 'no-cache');
    assert.deepEqual((passwordResponse.json() as { data: unknown }).data, { password: CLEAR_PASSWORD });
    assert.equal(passwordMock.mock.callCount(), 1);

    const defaultDetail = await app.inject({ method: 'GET', url: '/admin/emails/7', headers: { authorization: 'Bearer valid' } });
    assert.equal(defaultDetail.statusCode, 200);
    const defaultData = (defaultDetail.json() as { data: Record<string, unknown> }).data;
    assert.equal('password' in defaultData, false);
    assert.equal('refreshToken' in defaultData, false);
    assert.equal(detailMock.mock.calls.at(-1)?.arguments[1], false);

    assert.equal((await app.inject({
        method: 'GET',
        url: '/admin/emails/7?secrets=true',
        headers: { authorization: 'Bearer valid' },
    })).statusCode, 403);
    assert.equal(detailMock.mock.callCount(), 1);

    const superDetail = await app.inject({ method: 'GET', url: '/admin/emails/7?secrets=true', headers: superHeaders });
    assert.equal(superDetail.statusCode, 200);
    assert.equal(superDetail.headers['cache-control'], 'no-store');
    assert.equal(superDetail.headers.pragma, 'no-cache');
    assert.equal(detailMock.mock.calls.at(-1)?.arguments[1], true);

    const auditEntry = lines
        .map((line) => JSON.parse(line) as Record<string, unknown>)
        .find((entry) => entry.action === 'email.password_reveal');
    assert.ok(auditEntry);
    assert.equal(auditEntry.actorId, 42);
    assert.equal(auditEntry.actorUsername, 'auditor');
    assert.equal(auditEntry.emailId, 7);
    assert.equal('password' in auditEntry, false);
    for (const secret of [CLEAR_PASSWORD, CIPHERTEXT_MARKER]) {
        assert.equal(lines.some((line) => line.includes(secret)), false);
    }

    await app.close();
});
