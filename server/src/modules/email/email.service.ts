import prisma from '../../lib/prisma.js';
import { encrypt, decrypt } from '../../lib/crypto.js';
import { AppError } from '../../plugins/error.js';
import { requiresReauthorization } from '../../lib/microsoft-oauth.js';
import type { Prisma } from '@prisma/client';
import type { CreateEmailInput, UpdateEmailInput, ListEmailInput, ImportEmailInput } from './email.schema.js';

type EmailStatus = 'ACTIVE' | 'ERROR' | 'DISABLED';

export function statusUpdateDecision(
    current: { status: EmailStatus; errorMessage: string | null },
    requested: { status: EmailStatus; errorMessage?: string | null },
): { status?: EmailStatus; errorMessage?: string | null } {
    // A reauthorization marker is sticky until a new refresh token is saved.
    // Successful downstream work may have used an already cached access token,
    // while an unrelated downstream failure must not hide the actionable cause.
    if (requiresReauthorization(current.errorMessage)) return {};
    return { status: requested.status, errorMessage: requested.errorMessage || null };
}

type EmailStatusDb = Pick<Prisma.TransactionClient, 'emailAccount'>;
type EmailPasswordDb = Pick<Prisma.TransactionClient, 'emailAccount'>;
export type EmailImportDb = Pick<Prisma.TransactionClient, 'emailAccount' | 'emailGroup'>;

export async function importEmailAccountsWithDb(
    db: EmailImportDb,
    input: ImportEmailInput,
): Promise<{ success: number; failed: number; errors: string[] }> {
    const { content, separator, groupId } = input;
    if (!separator) {
        throw new AppError('INVALID_SEPARATOR', 'Import separator cannot be empty', 400);
    }

    const lines = content
        .split(/\r?\n/)
        .map((line, index) => ({ value: line.trim(), lineNumber: index + 1 }))
        .filter(({ value }) => Boolean(value));

    if (groupId !== undefined) {
        const group = await db.emailGroup.findUnique({ where: { id: groupId } });
        if (!group) {
            throw new AppError('GROUP_NOT_FOUND', 'Email group not found', 404);
        }
    }

    let success = 0;
    let failed = 0;
    const errors: string[] = [];

    for (const { value: line, lineNumber } of lines) {
        try {
            const parts = line.split(separator);
            if (parts.length < 3) {
                throw new Error('Invalid format');
            }

            let email: string;
            let clientId: string;
            let refreshToken: string;
            let password: string | undefined;

            if (parts.length >= 5) {
                [email, clientId] = parts;
                refreshToken = parts[4];
            } else if (parts.length === 4) {
                email = parts[0];
                password = parts[1] || undefined;
                clientId = parts[2];
                refreshToken = parts[3];
            } else {
                [email, clientId, refreshToken] = parts;
            }
            if (!email || !clientId || !refreshToken) {
                throw new Error('Missing required fields');
            }

            const now = new Date();
            const encryptedToken = encrypt(refreshToken);
            const updateData: Prisma.EmailAccountUncheckedUpdateInput = {
                clientId,
                refreshToken: encryptedToken,
                tokenVersion: { increment: 1 },
                tokenRefreshedAt: now,
                errorMessage: null,
                status: 'ACTIVE',
            };
            const createData: Prisma.EmailAccountUncheckedCreateInput = {
                email,
                clientId,
                refreshToken: encryptedToken,
                status: 'ACTIVE',
            };

            if (password) {
                const encryptedPassword = encrypt(password);
                updateData.password = encryptedPassword;
                createData.password = encryptedPassword;
            }
            if (groupId !== undefined) {
                updateData.groupId = groupId;
                createData.groupId = groupId;
            }

            await db.emailAccount.upsert({
                where: { email },
                update: updateData,
                create: createData,
            });
            success += 1;
        } catch (err) {
            failed += 1;
            const message = err instanceof Error && ['Invalid format', 'Missing required fields'].includes(err.message)
                ? err.message
                : 'Unable to save email account';
            errors.push(`Line ${lineNumber}: ${message}`);
        }
    }

    return { success, failed, errors };
}

export async function getEmailPasswordWithDb(
    db: EmailPasswordDb,
    id: number,
): Promise<{ password: string | null }> {
    if (!Number.isSafeInteger(id) || id <= 0) {
        throw new AppError('INVALID_ID', 'Invalid email account ID', 400);
    }

    const email = await db.emailAccount.findUnique({
        where: { id },
        select: { password: true },
    });

    if (!email) {
        throw new AppError('NOT_FOUND', 'Email account not found', 404);
    }

    if (!email.password) {
        return { password: null };
    }

    try {
        return { password: decrypt(email.password) };
    } catch {
        throw new AppError('PASSWORD_UNAVAILABLE', 'Unable to retrieve email password', 500);
    }
}

export async function updateEmailStatusWithDb(
    db: EmailStatusDb,
    id: number,
    status: EmailStatus,
    errorMessage?: string | null,
    tokenVersion?: number,
): Promise<void> {
    const current = await db.emailAccount.findUnique({
        where: { id },
        select: { tokenVersion: true, status: true, errorMessage: true },
    });
    if (!current || current.status === 'DISABLED' || (tokenVersion !== undefined && tokenVersion !== current.tokenVersion)) return;

    if (requiresReauthorization(errorMessage)) {
        // Reauthorization is actionable account state, not a routine check result.
        // It may supersede a concurrent ordinary status/error change, but never a
        // credential rotation or a concurrently disabled account.
        await db.emailAccount.updateMany({
            where: {
                id,
                tokenVersion: current.tokenVersion,
                status: { not: 'DISABLED' },
            },
            data: {
                status: 'ERROR',
                errorMessage: errorMessage || null,
                lastCheckAt: new Date(),
            },
        });
        return;
    }

    const update = statusUpdateDecision(current, { status, errorMessage });
    await db.emailAccount.updateMany({
        where: {
            id,
            tokenVersion: current.tokenVersion,
            status: current.status,
            errorMessage: current.errorMessage,
        },
        data: {
            ...update,
            lastCheckAt: new Date(),
        },
    });
}

export const emailService = {
    /**
     * 获取邮箱列表
     */
    async list(input: ListEmailInput) {
        const { page, pageSize, status, keyword, groupId, groupName } = input;
        const skip = (page - 1) * pageSize;

        const where: Prisma.EmailAccountWhereInput = {};
        if (status) where.status = status;
        if (keyword) {
            where.email = { contains: keyword };
        }
        if (groupId) {
            where.groupId = groupId;
        } else if (groupName) {
            where.group = { name: groupName };
        }

        const [list, total] = await Promise.all([
            prisma.emailAccount.findMany({
                where,
                select: {
                    id: true,
                    email: true,
                    clientId: true,
                status: true,
                groupId: true,
                group: { select: { id: true, name: true, fetchStrategy: true } },
                lastCheckAt: true,
                tokenRefreshedAt: true,
                errorMessage: true,
                createdAt: true,
            },
                skip,
                take: pageSize,
                orderBy: { id: 'desc' },
            }),
            prisma.emailAccount.count({ where }),
        ]);

        return { list, total, page, pageSize };
    },

    /**
     * 获取邮箱详情
     */
    async getById(id: number, includeSecrets = false) {
        const email = await prisma.emailAccount.findUnique({
            where: { id },
            select: {
                id: true,
                email: true,
                clientId: true,
                password: !!includeSecrets,
                refreshToken: !!includeSecrets,
                tokenVersion: true,
                status: true,
                groupId: true,
                group: { select: { id: true, name: true, fetchStrategy: true } },
                lastCheckAt: true,
                tokenRefreshedAt: true,
                errorMessage: true,
                createdAt: true,
                updatedAt: true,
            },
        });

        if (!email) {
            throw new AppError('NOT_FOUND', 'Email account not found', 404);
        }

        // 解密敏感信息
        if (includeSecrets) {
            return {
                ...email,
                refreshToken: email.refreshToken ? decrypt(email.refreshToken) : email.refreshToken,
                password: email.password ? decrypt(email.password) : email.password,
            };
        }

        return email;
    },

    /**
     * 按需读取单个邮箱账户保存的密码。
     * 权限控制由管理端路由负责；此处只读取并解密密码字段。
     */
    async getPassword(id: number): Promise<{ password: string | null }> {
        return getEmailPasswordWithDb(prisma, id);
    },

    /**
     * 根据邮箱地址获取（用于外部 API）
     */
    async getByEmail(emailAddress: string) {
        const email = await prisma.emailAccount.findUnique({
            where: { email: emailAddress },
            select: {
                id: true,
                email: true,
                clientId: true,
                refreshToken: true,
                tokenVersion: true,
                password: true,
                status: true,
                groupId: true,
                group: {
                    select: {
                        fetchStrategy: true,
                    },
                },
            },
        });

        if (!email) {
            return null;
        }

        // 解密
        return {
            ...email,
            refreshToken: decrypt(email.refreshToken),
            password: email.password ? decrypt(email.password) : undefined,
            fetchStrategy: email.group?.fetchStrategy || 'GRAPH_FIRST',
        };
    },

    /**
     * 创建邮箱账户
     */
    async create(input: CreateEmailInput) {
        const { email, clientId, refreshToken, password, groupId } = input;

        const exists = await prisma.emailAccount.findUnique({ where: { email } });
        if (exists) {
            throw new AppError('DUPLICATE_EMAIL', 'Email already exists', 400);
        }

        const encryptedToken = encrypt(refreshToken);
        const encryptedPassword = password ? encrypt(password) : null;

        const account = await prisma.emailAccount.create({
            data: {
                email,
                clientId,
                refreshToken: encryptedToken,
                password: encryptedPassword,
                groupId: groupId || null,
            },
            select: {
                id: true,
                email: true,
                clientId: true,
                status: true,
                groupId: true,
                createdAt: true,
            },
        });

        return account;
    },

    /**
     * 更新邮箱账户
     */
    async update(id: number, input: UpdateEmailInput) {
        const exists = await prisma.emailAccount.findUnique({ where: { id } });
        if (!exists) {
            throw new AppError('NOT_FOUND', 'Email account not found', 404);
        }

        const { refreshToken, password, ...rest } = input;
        const updateData: Prisma.EmailAccountUpdateInput = { ...rest, tokenVersion: { increment: 1 } };

        // 加密 sensitive data
        if (refreshToken) {
            updateData.refreshToken = encrypt(refreshToken);
            updateData.tokenRefreshedAt = new Date();
            updateData.errorMessage = null;
            if (!rest.status) updateData.status = 'ACTIVE';
        }
        if (password) {
            updateData.password = encrypt(password);
        }

        const account = await prisma.emailAccount.update({
            where: { id },
            data: updateData,
            select: {
                id: true,
                email: true,
                clientId: true,
                status: true,
                updatedAt: true,
            },
        });

        return account;
    },

    /**
     * 更新邮箱状态
     */
    async updateStatus(id: number, status: EmailStatus, errorMessage?: string | null, tokenVersion?: number) {
        await updateEmailStatusWithDb(prisma, id, status, errorMessage, tokenVersion);
    },

    /**
     * 仅更新时间，不改动邮箱状态
     */
    async touchLastCheckAt(id: number) {
        await prisma.emailAccount.update({
            where: { id },
            data: {
                lastCheckAt: new Date(),
            },
        });
    },

    /**
     * 删除邮箱账户
     */
    async delete(id: number) {
        const exists = await prisma.emailAccount.findUnique({ where: { id } });
        if (!exists) {
            throw new AppError('NOT_FOUND', 'Email account not found', 404);
        }

        await prisma.emailAccount.delete({ where: { id } });
        return { success: true };
    },

    /**
     * 批量删除
     */
    async batchDelete(ids: number[]) {
        await prisma.emailAccount.deleteMany({
            where: { id: { in: ids } },
        });
        return { deleted: ids.length };
    },

    /**
     * 批量导入
     */
    async import(input: ImportEmailInput) {
        return importEmailAccountsWithDb(prisma, input);
    },

    /**
     * 导出
     */
    async export(ids?: number[], separator = '----', groupId?: number) {
        const where: Prisma.EmailAccountWhereInput = {};
        if (ids?.length) {
            where.id = { in: ids };
        }
        if (groupId !== undefined) {
            where.groupId = groupId;
        }

        const accounts = await prisma.emailAccount.findMany({
            where,
            select: {
                email: true,
                clientId: true,
                refreshToken: true,
            },
        });

        const lines = accounts.map((acc: { email: string; clientId: string; refreshToken: string }) => {
            const token = decrypt(acc.refreshToken);
            return `${acc.email}${separator}${acc.clientId}${separator}${token}`;
        });

        return lines.join('\n');
    },

    /**
     * 获取统计
     */
    async getStats() {
        const [total, active, error] = await Promise.all([
            prisma.emailAccount.count(),
            prisma.emailAccount.count({ where: { status: 'ACTIVE' } }),
            prisma.emailAccount.count({ where: { status: 'ERROR' } }),
        ]);

        return { total, active, error };
    },
};
