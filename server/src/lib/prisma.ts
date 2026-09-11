import { PrismaClient } from '@prisma/client';
import { env } from '../config/env.js';

const globalForPrisma = globalThis as unknown as {
    prisma: PrismaClient | undefined;
};

export const prisma = globalForPrisma.prisma ?? new PrismaClient({
    // Prisma's automatic error output can include query arguments (encrypted
    // tokens and device codes). Callers report normalized operation failures.
    log: env.NODE_ENV === 'development' ? ['warn'] : [],
});

if (env.NODE_ENV !== 'production') {
    globalForPrisma.prisma = prisma;
}

export default prisma;
