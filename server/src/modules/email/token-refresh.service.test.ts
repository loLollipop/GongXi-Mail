import assert from 'node:assert/strict';
import test from 'node:test';
import type { Prisma } from '@prisma/client';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
process.env.JWT_SECRET = 'test-jwt-secret-for-token-refresh-0000';

const { persistTokenRefreshFailure } = await import('./token-refresh.service.js');

type EmailStatusDb = Pick<Prisma.TransactionClient, 'emailAccount'>;

void test('an ordinary refresh failure cannot overwrite the refresh service reauthorization marker', async () => {
    const state: { tokenVersion: number; errorMessage: string | null; status: 'ACTIVE' | 'ERROR' | 'DISABLED' } = {
        tokenVersion: 4,
        errorMessage: null,
        status: 'ACTIVE',
    };
    const queries: Prisma.EmailAccountUpdateManyArgs[] = [];
    const db = { emailAccount: {
        findUnique: async () => ({ ...state }),
        updateMany: async (input: Prisma.EmailAccountUpdateManyArgs) => {
            queries.push(input);
            if (state.tokenVersion !== input.where?.tokenVersion || state.status === 'DISABLED') return { count: 0 };
            if (typeof input.where?.status === 'string' && state.status !== input.where.status) return { count: 0 };
            if ('errorMessage' in (input.where ?? {}) && state.errorMessage !== input.where?.errorMessage) return { count: 0 };
            state.errorMessage = input.data.errorMessage as string;
            if (input.data.status) state.status = input.data.status as 'ACTIVE' | 'ERROR' | 'DISABLED';
            return { count: 1 };
        },
    } } as unknown as EmailStatusDb;

    await persistTokenRefreshFailure(7, 4, 'invalid_grant: AADSTS65001: REAUTHORIZATION_REQUIRED', db);
    const marker = state.errorMessage;
    await persistTokenRefreshFailure(7, 4, 'Exception: Token refresh network or persistence failure', db);

    assert.match(marker ?? '', /^Token refresh: .*AADSTS65001/);
    assert.equal(state.errorMessage, marker);
    assert.equal(queries.length, 1, 'the ordinary failure is discarded before issuing a write');
    assert.deepEqual(queries[0]?.where, {
        id: 7,
        tokenVersion: 4,
        status: { not: 'DISABLED' },
    });
});

void test('token refresh preserves unrelated errors while reauthorization remains a priority write', async () => {
    const state: { errorMessage: string | null; status: 'ACTIVE' | 'ERROR' } = {
        errorMessage: 'Mailbox quota exceeded',
        status: 'ERROR',
    };
    const queries: Prisma.EmailAccountUpdateManyArgs[] = [];
    const db = { emailAccount: {
        findUnique: async () => ({ tokenVersion: 4, ...state }),
        updateMany: async (input: Prisma.EmailAccountUpdateManyArgs) => {
            queries.push(input);
            state.errorMessage = input.data.errorMessage as string;
            if (input.data.status) state.status = input.data.status as 'ACTIVE' | 'ERROR';
            return { count: 1 };
        },
    } } as unknown as EmailStatusDb;

    await persistTokenRefreshFailure(7, 4, 'HTTP 503', db);
    assert.equal(state.errorMessage, 'Mailbox quota exceeded');
    await persistTokenRefreshFailure(7, 4, 'invalid_grant: AADSTS65001: REAUTHORIZATION_REQUIRED', db);

    assert.equal(queries.length, 1, 'the unrelated existing error is not overwritten');
    assert.deepEqual(queries[0]?.where, {
        id: 7,
        tokenVersion: 4,
        status: { not: 'DISABLED' },
    });
    assert.equal(queries[0]?.data.status, 'ERROR');
    assert.match(queries[0]?.data.errorMessage as string, /AADSTS65001/);
    assert.match(state.errorMessage as string, /AADSTS65001/);
});
