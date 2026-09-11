import assert from 'node:assert/strict';
import test from 'node:test';
import type { Prisma } from '@prisma/client';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.JWT_SECRET = 'test-jwt-secret-for-reauthorization-0000';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
const prisma = { emailAccount: { updateMany: async () => ({ count: 0 }) } } as unknown as Pick<Prisma.TransactionClient, 'emailAccount'>;
const { persistRotatedToken } = await import('./rotated-token.js');
const { decrypt } = await import('./crypto.js');

void test('rotation encrypts, checks the old version and advances in-flight credentials only after winning CAS', async (t) => {
    let query: Prisma.EmailAccountUpdateManyArgs | undefined;
    t.mock.method(prisma.emailAccount, 'updateMany', async (input: Prisma.EmailAccountUpdateManyArgs) => {
        query = input;
        return { count: 1 };
    });
    const credentials = { id: 7, tokenVersion: 4, refreshToken: 'old' };
    assert.equal(await persistRotatedToken(credentials, 'rotated', prisma), true);
    assert.deepEqual(query?.where, { id: 7, tokenVersion: 4, status: { not: 'DISABLED' } });
    assert.deepEqual(query?.data.tokenVersion, { increment: 1 });
    assert.equal(decrypt(query?.data.refreshToken as string), 'rotated');
    assert.deepEqual(credentials, { id: 7, tokenVersion: 5, refreshToken: 'rotated' });
});

void test('lost CAS cannot overwrite a newer token or mutate fallback credentials', async (t) => {
    const update = t.mock.method(prisma.emailAccount, 'updateMany', async () => ({ count: 0 }));
    const credentials = { id: 7, tokenVersion: 4, refreshToken: 'old' };
    assert.equal(await persistRotatedToken(credentials, 'stale', prisma), false);
    assert.deepEqual(credentials, { id: 7, tokenVersion: 4, refreshToken: 'old' });
    assert.equal(await persistRotatedToken(credentials, undefined, prisma), true);
    assert.equal(update.mock.callCount(), 1);
});
