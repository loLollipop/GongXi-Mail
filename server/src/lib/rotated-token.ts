import prisma from './prisma.js';
import { encrypt } from './crypto.js';
import type { Prisma } from '@prisma/client';

export interface TokenCredentials {
    id: number;
    tokenVersion: number;
    refreshToken: string;
}

/** Update the in-flight credentials only when this exact version won the CAS.
 * Versioned access-token cache keys make older cached tokens unreachable.
 * Shared by mail access and scheduled refresh; no dependency on mail services.
 */
export async function persistRotatedToken(
    credentials: TokenCredentials,
    refreshToken: string | undefined,
    db: Pick<Prisma.TransactionClient, 'emailAccount'> = prisma,
    extra: Pick<Prisma.EmailAccountUpdateManyMutationInput, 'errorMessage' | 'status'> = {},
): Promise<boolean> {
    if (!refreshToken) return true;
    const result = await db.emailAccount.updateMany({
        where: { id: credentials.id, tokenVersion: credentials.tokenVersion, status: { not: 'DISABLED' } },
        data: {
            ...extra,
            refreshToken: encrypt(refreshToken),
            tokenRefreshedAt: new Date(),
            tokenVersion: { increment: 1 },
        },
    });
    if (result.count !== 1) return false;
    credentials.refreshToken = refreshToken;
    credentials.tokenVersion += 1;
    return true;
}
