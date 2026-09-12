import assert from 'node:assert/strict';
import test, { type TestContext } from 'node:test';
import Fastify from 'fastify';
import type { EmailReauthorization, Prisma, PrismaClient } from '@prisma/client';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.JWT_SECRET = 'test-jwt-secret-for-reauthorization-0000';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
// A plain injected double avoids mutating Prisma's proxy delegates or opening
// a database connection. Query predicates remain asserted below.
const prisma = {
    emailReauthorization: { findUnique: async () => null, updateMany: async () => ({ count: 0 }),
        create: async () => null, update: async () => null, deleteMany: async () => ({ count: 0 }) },
    emailAccount: { findUnique: async () => null, findMany: async () => [], updateMany: async () => ({ count: 0 }) },
    $transaction: async () => undefined,
} as unknown as PrismaClient;
const { encrypt, decrypt } = await import('../../lib/crypto.js');
const { createReauthorizationService, hasGraphMailReadWriteScope, identityMatches, pollingOutcome,
    publicSession, reauthorizationService, DEVICE_SCOPES } = await import('./reauthorization.service.js');
const { default: routes } = await import('./reauthorization.routes.js');
const { AppError } = await import('../../plugins/error.js');

function fixture(): EmailReauthorization {
    return {
        id: '12345678-1234-4123-8123-123456789012', emailId: 7, activeEmailId: 7,
        email: 'target@outlook.com', clientId: 'client', tokenVersion: 4, createdBy: 1,
        status: 'PENDING', deviceCode: encrypt('secret-device-code'), userCode: 'ABCD-EFGH',
        verificationUri: 'https://microsoft.com/devicelogin', verificationUriComplete: null,
        expiresAt: new Date(Date.now() + 600000), interval: 5, nextPollAt: new Date(Date.now() - 1000),
        pollLeaseUntil: null, pollClaim: null, errorCode: null, errorMessage: null,
        authorizedEmail: null, createdAt: new Date(), updatedAt: new Date(), completedAt: null,
    };
}

function mockStore(t: TestContext, initial = fixture(), savedCount = 1) {
    let row = initial;
    const writes: Prisma.EmailAccountUpdateManyArgs[] = [];
    t.mock.method(prisma.emailReauthorization, 'findUnique', async () => ({ ...row }));
    t.mock.method(prisma.emailReauthorization, 'deleteMany', async () => ({ count: 0 }));
    t.mock.method(prisma.emailReauthorization, 'updateMany', async (input: Prisma.EmailReauthorizationUpdateManyArgs) => {
        const where = input.where;
        // Cleanup's global queries are irrelevant to an unexpired fixture.
        if (!where?.id) return { count: 0 };
        if (typeof where.status === 'string' && row.status !== where.status) return { count: 0 };
        if (where.status && typeof where.status === 'object' && Array.isArray(where.status.in) && !where.status.in.includes(row.status)) return { count: 0 };
        if (where.pollClaim && row.pollClaim !== where.pollClaim) return { count: 0 };
        row = { ...row, ...input.data } as EmailReauthorization;
        return { count: 1 };
    });
    t.mock.method(prisma.emailReauthorization, 'update', async (input: Prisma.EmailReauthorizationUpdateArgs) => {
        row = { ...row, ...input.data } as EmailReauthorization;
        return { ...row };
    });
    t.mock.method(prisma.emailAccount, 'updateMany', async (input: Prisma.EmailAccountUpdateManyArgs) => {
        writes.push(input);
        return { count: savedCount };
    });
    t.mock.method(prisma, '$transaction', async (fn: (tx: Prisma.TransactionClient) => Promise<void>) => fn(prisma));
    return { writes, row: () => row };
}

void test('identity check fails closed and scopes use one Graph resource', () => {
    assert.equal(identityMatches('Target@outlook.com', { mail: 'target@OUTLOOK.com' }).matches, true);
    assert.equal(identityMatches('target@outlook.com', { userPrincipalName: 'target@outlook.com' }).matches, true);
    assert.equal(identityMatches('target@outlook.com', { mail: 'attacker@outlook.com', displayName: 'target@outlook.com' }).matches, false);
    assert.equal(identityMatches('target@outlook.com', {}).matches, false);
    assert.match(DEVICE_SCOPES, /offline_access/);
    assert.match(DEVICE_SCOPES, /graph\.microsoft\.com\/Mail\.ReadWrite/);
    assert.doesNotMatch(DEVICE_SCOPES, /graph\.microsoft\.com\/Mail\.Read(?:\s|$)/);
    assert.doesNotMatch(DEVICE_SCOPES, /outlook\.office/);
    assert.equal(hasGraphMailReadWriteScope('openid MAIL.READWRITE'), true);
    assert.equal(hasGraphMailReadWriteScope('https://graph.microsoft.com/Mail.ReadWrite User.Read'), true);
    assert.equal(hasGraphMailReadWriteScope('Mail.Read User.Read'), false);
    assert.equal(hasGraphMailReadWriteScope(undefined), false);
});

void test('public DTO explicitly excludes device codes, claims, client IDs and token versions', () => {
    const value = publicSession(fixture());
    for (const key of ['deviceCode', 'device_code', 'access_token', 'refresh_token', 'password', 'clientId', 'tokenVersion', 'pollClaim']) assert.equal(key in value, false, key);
    assert.doesNotMatch(JSON.stringify(value), /secret-device-code/);
});

void test('polling outcomes enforce slow_down and terminal rejection/expiry', () => {
    assert.deepEqual(pollingOutcome('authorization_pending', 5), { status: 'PENDING', interval: 5 });
    assert.deepEqual(pollingOutcome('slow_down', 5), { status: 'PENDING', interval: 10 });
    assert.equal(pollingOutcome('authorization_declined', 5).status, 'DECLINED');
    assert.equal(pollingOutcome('expired_token', 5).status, 'EXPIRED');
});

void test('an early poll never contacts Microsoft', async (t) => {
    const initial = fixture();
    initial.nextPollAt = new Date(Date.now() + 30000);
    mockStore(t, initial);
    let calls = 0;
    const service = createReauthorizationService(prisma, async () => { calls++; return new Response(); });
    assert.equal((await service.poll(initial.id)).status, 'PENDING');
    assert.equal(calls, 0);
});

void test('start reserves the unique account slot before contacting Microsoft and encrypts the device code', async (t) => {
    const state = mockStore(t);
    let stored: Prisma.EmailReauthorizationCreateArgs | undefined;
    let created = false;
    t.mock.method(prisma.emailAccount, 'findUnique', async () => ({ id: 7, email: 'target@outlook.com', clientId: 'client', tokenVersion: 4, status: 'ERROR', errorMessage: 'AADSTS70000', group: { fetchStrategy: 'GRAPH_FIRST' } }));
    t.mock.method(prisma.emailReauthorization, 'findUnique', async (input: Prisma.EmailReauthorizationFindUniqueArgs) => input.where.activeEmailId && !created ? null : state.row());
    t.mock.method(prisma.emailReauthorization, 'create', async (input: Prisma.EmailReauthorizationCreateArgs) => {
        stored = input;
        created = true;
        Object.assign(state.row(), input.data, { status: 'STARTING' });
        return state.row();
    });
    let calls = 0;
    const service = createReauthorizationService(prisma, async (url, options) => {
        calls++;
        assert.equal(created, true);
        assert.match(url, /consumers\/oauth2\/v2.0\/devicecode$/);
        assert.equal(new URLSearchParams(options?.body as string).get('scope'), DEVICE_SCOPES);
        return Response.json({ device_code: 'new-secret-device', user_code: 'ABCD-EFGH', verification_uri: 'https://microsoft.com/devicelogin', expires_in: 900, interval: 5 });
    });
    const result = await service.start(7, 1);
    assert.equal(stored?.data.activeEmailId, 7);
    assert.equal(stored?.data.tokenVersion, 4);
    assert.equal(result.status, 'PENDING');
    assert.equal(decrypt(state.row().deviceCode!), 'new-secret-device');
    assert.doesNotMatch(JSON.stringify(result), /new-secret-device/);
    assert.equal((await service.start(7, 1)).sessionId, result.sessionId);
    assert.equal(calls, 1);
});

void test('disabled and unrelated errors cannot start authorization', async (t) => {
    mockStore(t);
    let calls = 0;
    const service = createReauthorizationService(prisma, async () => { calls++; return new Response(); });
    const find = t.mock.method(prisma.emailAccount, 'findUnique', async () => ({ status: 'DISABLED', errorMessage: 'AADSTS70000' }));
    await assert.rejects(service.start(7, 1), { code: 'ACCOUNT_UNAVAILABLE' });
    find.mock.mockImplementation(async () => ({ status: 'ERROR', errorMessage: 'HTTP 503' }));
    await assert.rejects(service.start(7, 1), { code: 'REAUTHORIZATION_NOT_REQUIRED' });
    assert.equal(calls, 0);
});

void test('IMAP_ONLY accounts cannot start automatic Graph reauthorization', async (t) => {
    mockStore(t);
    let calls = 0;
    t.mock.method(prisma.emailAccount, 'findUnique', async () => ({
        id: 7, email: 'target@outlook.com', clientId: 'client', tokenVersion: 4,
        status: 'ERROR', errorMessage: 'AADSTS65001', group: { fetchStrategy: 'IMAP_ONLY' },
    }));
    const service = createReauthorizationService(prisma, async () => { calls++; return new Response(); });
    await assert.rejects(service.start(7, 1), { code: 'REAUTHORIZATION_STRATEGY_UNSUPPORTED' });
    assert.equal(calls, 0);
});

void test('candidates only include reauthorization errors and expose recoverable safe sessions', async (t) => {
    mockStore(t);
    t.mock.method(prisma.emailAccount, 'findMany', async () => [
        { id: 7, email: 'target@outlook.com', errorMessage: 'AADSTS70000 secret', group: { name: 'Graph', fetchStrategy: 'GRAPH_FIRST' }, reauthorizations: [fixture()] },
        { id: 9, email: 'imap@outlook.com', errorMessage: 'AADSTS65001', group: { name: 'IMAP', fetchStrategy: 'IMAP_ONLY' }, reauthorizations: [] },
        { id: 8, email: 'network@outlook.com', errorMessage: 'HTTP 503', group: null, reauthorizations: [] },
    ]);
    const result = await createReauthorizationService(prisma).candidates();
    assert.equal(result.length, 2);
    assert.equal(result[0].activeSession?.sessionId, fixture().id);
    assert.equal(result[0].supported, true);
    assert.equal(result[0].unsupportedReason, null);
    assert.equal(result[1].supported, false);
    assert.match(result[1].unsupportedReason ?? '', /仅 IMAP/);
    assert.equal(result.some((candidate) => 'password' in candidate || !!candidate.activeSession && 'password' in candidate.activeSession), false);
    assert.doesNotMatch(JSON.stringify(result), /secret/);
});

void test('concurrent polls issue one Microsoft request and slow_down updates the persisted schedule', async (t) => {
    const state = mockStore(t);
    let calls = 0;
    const service = createReauthorizationService(prisma, async () => {
        calls++;
        return Response.json({ error: 'slow_down' }, { status: 400 });
    });
    await Promise.all([service.poll(state.row().id), service.poll(state.row().id)]);
    assert.equal(calls, 1);
    assert.equal(state.row().status, 'PENDING');
    assert.equal(state.row().interval, 10);
    assert.ok(state.row().nextPollAt.getTime() >= Date.now() + 9000);
});

void test('network failure backs off without disclosing request secrets', async (t) => {
    const state = mockStore(t);
    const service = createReauthorizationService(prisma, async () => { throw new Error('secret-device-code'); });
    const result = await service.poll(state.row().id);
    assert.equal(result.status, 'PENDING');
    assert.equal(result.interval, 10);
    assert.doesNotMatch(JSON.stringify(result), /secret-device-code/);
});

for (const scenario of ['success', 'mismatch', 'missing_token', 'missing_scope', 'cas_lost', 'cancel_in_flight'] as const) {
    void test(`redemption ${scenario} preserves account/session invariants`, async (t) => {
        const state = mockStore(t, fixture(), scenario === 'cas_lost' ? 0 : 1);
        let calls = 0;
        const service = createReauthorizationService(prisma, async (url) => {
            calls++;
            if (url.includes('/token')) return Response.json({ access_token: 'secret-access',
                ...(scenario === 'missing_token' ? {} : { refresh_token: 'secret-refresh' }),
                ...(scenario === 'missing_scope' ? {} : { scope: 'openid https://graph.microsoft.com/Mail.ReadWrite' }) });
            if (scenario === 'cancel_in_flight') await service.cancel(state.row().id);
            return Response.json({ mail: scenario === 'mismatch' ? 'other@outlook.com' : 'target@outlook.com' });
        });
        const result = await service.poll(state.row().id);
        assert.doesNotMatch(JSON.stringify(result), /secret-access|secret-refresh|secret-device-code/);
        assert.equal(state.row().deviceCode, null);
        assert.equal(state.row().activeEmailId, null);
        if (scenario === 'success' || scenario === 'cas_lost') {
            assert.equal(state.writes.length, 1);
            assert.equal(state.writes[0].where?.tokenVersion, 4);
            assert.equal(decrypt(state.writes[0].data.refreshToken as string), 'secret-refresh');
            assert.equal(result.status, scenario === 'success' ? 'SUCCEEDED' : 'FAILED');
            if (scenario === 'cas_lost') assert.equal(result.errorCode, 'ACCOUNT_CHANGED');
        } else {
            assert.equal(state.writes.length, 0);
            assert.equal(result.status, scenario === 'mismatch' ? 'MISMATCH' : scenario === 'cancel_in_flight' ? 'CANCELLED' : 'FAILED');
            if (scenario === 'missing_scope') assert.equal(result.errorCode, 'GRAPH_MAIL_READWRITE_SCOPE_MISSING');
        }
        assert.equal(calls, scenario === 'missing_token' || scenario === 'missing_scope' ? 1 : 2);
    });
}

void test('all six endpoints require JWT then SUPER_ADMIN; validation runs after authentication', async (t) => {
    const app = Fastify();
    app.decorate('authenticateJwt', async (request) => {
        if (request.headers.authorization !== 'Bearer valid') throw new AppError('UNAUTHORIZED', 'Authentication required', 401);
        request.user = { id: 1, username: 'test', role: request.headers['x-test-role'] === 'super' ? 'SUPER_ADMIN' : 'ADMIN' };
    });
    app.decorate('requireSuperAdmin', async (request) => {
        if (request.user?.role !== 'SUPER_ADMIN') throw new AppError('FORBIDDEN', 'Super admin access required', 403);
    });
    t.mock.method(reauthorizationService, 'candidates', async () => []);
    await app.register(routes, { prefix: '/admin/email-reauthorizations' });
    const id = fixture().id;
    for (const [method, path] of [['GET', '/candidates'], ['POST', '/start'], ['GET', `/${id}`], ['POST', `/${id}/poll`], ['POST', `/${id}/cancel`]] as const) {
        for (const [headers, status] of [[{}, 401], [{ authorization: 'Bearer valid' }, 403]] as const) {
            const response = await app.inject({ method, url: `/admin/email-reauthorizations${path}`, headers });
            assert.equal(response.statusCode, status, `${method} ${path}`);
        }
    }
    const headers = { authorization: 'Bearer valid', 'x-test-role': 'super' };
    assert.equal((await app.inject({ method: 'GET', url: '/admin/email-reauthorizations/candidates', headers })).statusCode, 200);
    assert.equal((await app.inject({ method: 'POST', url: '/admin/email-reauthorizations/start', headers, payload: { emailId: -1 } })).statusCode, 400);
    assert.equal((await app.inject({ method: 'GET', url: '/admin/email-reauthorizations/not-a-uuid', headers })).statusCode, 400);
    await app.close();
});
