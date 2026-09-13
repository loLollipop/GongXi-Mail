import assert from 'node:assert/strict';
import test from 'node:test';
import type { Prisma } from '@prisma/client';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
process.env.JWT_SECRET = 'test-jwt-secret-for-email-status-0000';

const { decrypt } = await import('../../lib/crypto.js');
const { importEmailAccountsWithDb, statusUpdateDecision, updateEmailStatusWithDb } = await import('./email.service.js');

type EmailStatusDb = Pick<Prisma.TransactionClient, 'emailAccount'>;
type EmailImportDb = Pick<Prisma.TransactionClient, 'emailAccount' | 'emailGroup'>;

void test('email import parses CRLF and four-column rows and uses upsert payloads', async () => {
    const upserts: Prisma.EmailAccountUpsertArgs[] = [];
    let groupLookups = 0;
    const db = {
        emailGroup: {
            findUnique: async () => {
                groupLookups += 1;
                return { id: 7 };
            },
        },
        emailAccount: {
            upsert: async (input: Prisma.EmailAccountUpsertArgs) => {
                upserts.push(input);
                return {};
            },
        },
    } as unknown as EmailImportDb;

    const result = await importEmailAccountsWithDb(db, {
        content: 'one@example.com---- secret ----client-1----refresh-1\r\n\r\nbroken\r\ntwo@example.com----client-2----refresh-2\r\nsix@example.com----client-6----uuid----info----refresh-6',
        separator: '----',
        groupId: 7,
    });

    assert.equal(groupLookups, 1);
    assert.deepEqual({ success: result.success, failed: result.failed }, { success: 3, failed: 1 });
    assert.match(result.errors[0] || '', /^Line 3: Invalid format$/);
    assert.equal(upserts.length, 3);

    const first = upserts[0];
    assert.deepEqual(first.where, { email: 'one@example.com' });
    assert.equal(first.update.clientId, 'client-1');
    assert.deepEqual(first.update.tokenVersion, { increment: 1 });
    assert.equal(first.update.status, 'ACTIVE');
    assert.equal(first.update.errorMessage, null);
    assert.equal(first.update.groupId, 7);
    assert.equal(decrypt(first.update.refreshToken as string), 'refresh-1');
    assert.equal(decrypt(first.update.password as string), ' secret ');
    assert.equal(first.create.groupId, 7);
    assert.equal(decrypt(first.create.refreshToken), 'refresh-1');
    assert.equal(decrypt(first.create.password as string), ' secret ');
    assert.ok(first.update.tokenRefreshedAt instanceof Date);
    assert.equal(Object.hasOwn(first.create, 'tokenRefreshedAt'), false);

    const second = upserts[1];
    assert.equal(Object.hasOwn(second.update, 'password'), false);
    assert.equal(Object.hasOwn(second.create, 'password'), false);

    const third = upserts[2];
    assert.equal(third.update.clientId, 'client-6');
    assert.equal(decrypt(third.update.refreshToken as string), 'refresh-6');
});

void test('email import leaves optional password and group unchanged when omitted', async () => {
    let upsert: Prisma.EmailAccountUpsertArgs | undefined;
    const db = {
        emailGroup: { findUnique: async () => null },
        emailAccount: {
            upsert: async (input: Prisma.EmailAccountUpsertArgs) => {
                upsert = input;
                return {};
            },
        },
    } as unknown as EmailImportDb;

    const result = await importEmailAccountsWithDb(db, {
        content: 'three@example.com----client-3----refresh-3',
        separator: '----',
    });

    assert.deepEqual(result, { success: 1, failed: 0, errors: [] });
    assert.ok(upsert);
    assert.equal(Object.hasOwn(upsert.update, 'password'), false);
    assert.equal(Object.hasOwn(upsert.update, 'groupId'), false);
    assert.equal(Object.hasOwn(upsert.create, 'password'), false);
    assert.equal(Object.hasOwn(upsert.create, 'groupId'), false);
});

void test('email import errors never echo credentials when the separator is wrong', async () => {
    const db = {
        emailGroup: { findUnique: async () => null },
        emailAccount: { upsert: async () => ({}) },
    } as unknown as EmailImportDb;
    const password = 'do-not-return-this-password';
    const refreshToken = 'do-not-return-this-refresh-token';

    const result = await importEmailAccountsWithDb(db, {
        content: `four@example.com----${password}----client-4----${refreshToken}`,
        separator: '|',
    });

    assert.deepEqual({ success: result.success, failed: result.failed }, { success: 0, failed: 1 });
    assert.match(result.errors[0] || '', /^Line 1: Invalid format$/);
    assert.equal(result.errors.join('\n').includes(password), false);
    assert.equal(result.errors.join('\n').includes(refreshToken), false);
});

void test('email import sanitizes database errors before returning them', async () => {
    const secret = 'database-error-must-not-return-this-token';
    const db = {
        emailGroup: { findUnique: async () => null },
        emailAccount: {
            upsert: async () => {
                throw new Error(`database rejected ${secret}`);
            },
        },
    } as unknown as EmailImportDb;

    const result = await importEmailAccountsWithDb(db, {
        content: 'five@example.com----client-5----refresh-5',
        separator: '----',
    });

    assert.deepEqual(result, {
        success: 0,
        failed: 1,
        errors: ['Line 1: Unable to save email account'],
    });
    assert.equal(result.errors.join('\n').includes(secret), false);
});

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
