import assert from 'node:assert/strict';
import test from 'node:test';
import type { Prisma } from '@prisma/client';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
process.env.JWT_SECRET = 'test-jwt-secret-for-email-status-0000';

const { statusUpdateDecision, updateEmailStatusWithDb } = await import('./email.service.js');

type EmailStatusDb = Pick<Prisma.TransactionClient, 'emailAccount'>;

void test('status updates preserve an existing reauthorization marker', () => {
    const current = { status: 'ERROR' as const, errorMessage: 'invalid_grant: AADSTS65001: REAUTHORIZATION_REQUIRED' };
    assert.deepEqual(statusUpdateDecision(current, { status: 'ACTIVE' }), {});
    assert.deepEqual(statusUpdateDecision(current, { status: 'ERROR', errorMessage: 'HTTP 503' }), {});
});

void test('status updates retain the previous behavior for non-reauthorization errors', () => {
    assert.deepEqual(
        statusUpdateDecision({ status: 'ERROR', errorMessage: 'HTTP 503' }, { status: 'ACTIVE' }),
        { status: 'ACTIVE', errorMessage: null },
    );
    assert.deepEqual(
        statusUpdateDecision({ status: 'ACTIVE', errorMessage: null }, { status: 'ERROR', errorMessage: 'IMAP failed' }),
        { status: 'ERROR', errorMessage: 'IMAP failed' },
    );
});

void test('a newly discovered reauthorization error wins an ordinary status/error race', async () => {
    const state: { tokenVersion: number; status: 'ACTIVE' | 'ERROR' | 'DISABLED'; errorMessage: string | null } = {
        tokenVersion: 4,
        status: 'ACTIVE',
        errorMessage: null,
    };
    const db = { emailAccount: {
        findUnique: async () => {
            const snapshot = { ...state };
            state.status = 'ERROR';
            state.errorMessage = 'HTTP 503';
            return snapshot;
        },
        updateMany: async (input: Prisma.EmailAccountUpdateManyArgs) => {
            assert.deepEqual(input.where, {
                id: 7,
                tokenVersion: 4,
                status: { not: 'DISABLED' },
            });
            if (state.tokenVersion === 4 && state.status !== 'DISABLED') {
                state.status = input.data.status as 'ACTIVE' | 'ERROR' | 'DISABLED';
                state.errorMessage = input.data.errorMessage as string | null;
                return { count: 1 };
            }
            return { count: 0 };
        },
    } } as unknown as EmailStatusDb;

    await updateEmailStatusWithDb(db, 7, 'ERROR', 'invalid_grant: AADSTS65001: REAUTHORIZATION_REQUIRED', 4);

    assert.equal(state.status, 'ERROR');
    assert.equal(state.errorMessage, 'invalid_grant: AADSTS65001: REAUTHORIZATION_REQUIRED');
});

void test('a reauthorization write cannot affect a rotated or concurrently disabled account', async () => {
    const updates: Prisma.EmailAccountUpdateManyArgs[] = [];
    const db = { emailAccount: {
        findUnique: async () => ({
            tokenVersion: 5,
            status: 'ACTIVE' as const,
            errorMessage: null,
        }),
        updateMany: async (input: Prisma.EmailAccountUpdateManyArgs) => {
            updates.push(input);
            return { count: 0 };
        },
    } } as unknown as EmailStatusDb;

    await updateEmailStatusWithDb(db, 7, 'ERROR', 'AADSTS65001: REAUTHORIZATION_REQUIRED', 4);
    assert.equal(updates.length, 0, 'a stale credential snapshot is rejected before writing');

    await updateEmailStatusWithDb(db, 7, 'ERROR', 'AADSTS65001: REAUTHORIZATION_REQUIRED', 5);
    assert.deepEqual(updates[0]?.where, {
        id: 7,
        tokenVersion: 5,
        status: { not: 'DISABLED' },
    });
});

void test('an ordinary success cannot clear a reauthorization marker written after its read', async () => {
    const state: { tokenVersion: number; status: 'ACTIVE' | 'ERROR' | 'DISABLED'; errorMessage: string | null } = {
        tokenVersion: 4,
        status: 'ACTIVE',
        errorMessage: null,
    };
    const db = { emailAccount: {
        findUnique: async () => {
            const snapshot = { ...state };
            state.status = 'ERROR';
            state.errorMessage = 'AADSTS65001: REAUTHORIZATION_REQUIRED';
            return snapshot;
        },
        updateMany: async (input: Prisma.EmailAccountUpdateManyArgs) => {
            assert.deepEqual(input.where, {
                id: 7,
                tokenVersion: 4,
                status: 'ACTIVE',
                errorMessage: null,
            });
            return { count: 0 };
        },
    } } as unknown as EmailStatusDb;

    await updateEmailStatusWithDb(db, 7, 'ACTIVE', undefined, 4);

    assert.equal(state.status, 'ERROR');
    assert.equal(state.errorMessage, 'AADSTS65001: REAUTHORIZATION_REQUIRED');
});
