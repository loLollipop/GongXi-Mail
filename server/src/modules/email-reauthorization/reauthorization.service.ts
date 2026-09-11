import { randomUUID } from 'node:crypto';
import { z } from 'zod';
import type { EmailReauthorization, PrismaClient, ReauthorizationStatus } from '@prisma/client';
import prisma from '../../lib/prisma.js';
import { encrypt, decrypt } from '../../lib/crypto.js';
import { proxyFetch } from '../../lib/proxy.js';
import { readMicrosoftError, requiresReauthorization } from '../../lib/microsoft-oauth.js';
import { AppError } from '../../plugins/error.js';

export const DEVICE_SCOPES = 'openid profile offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite';
export const IMAP_ONLY_UNSUPPORTED_REASON = '仅 IMAP 分组不支持此中心自动恢复；请改为可回退到 Graph 的策略后重试，或另行完成 IMAP 授权';
const MICROSOFT = 'https://login.microsoftonline.com/consumers/oauth2/v2.0';
const ACTIVE: ReauthorizationStatus[] = ['STARTING', 'PENDING', 'POLLING'];
const terminalData = { activeEmailId: null, deviceCode: null, userCode: null,
    verificationUriComplete: null, pollClaim: null, pollLeaseUntil: null };

const verificationUrl = z.string().url().refine((value) => {
    const url = new URL(value);
    return url.protocol === 'https:' && !url.username && !url.password &&
        ['microsoft.com', 'www.microsoft.com', 'login.microsoftonline.com', 'login.live.com'].includes(url.hostname);
});
const deviceResponse = z.object({
    device_code: z.string().min(1), user_code: z.string().min(1).max(100),
    verification_uri: verificationUrl, verification_uri_complete: verificationUrl.optional(),
    expires_in: z.number().int().positive().max(3600),
    interval: z.number().int().positive().max(300).default(5),
});
const tokenResponse = z.object({
    access_token: z.string().min(1),
    refresh_token: z.string().min(1),
    scope: z.string().optional(),
});

export function publicSession(session: EmailReauthorization) {
    const serverTime = new Date();
    return {
        sessionId: session.id, emailId: session.emailId, email: session.email,
        userCode: session.userCode, verificationUri: session.verificationUri,
        verificationUriComplete: session.verificationUriComplete,
        expiresAt: session.expiresAt, interval: session.interval,
        nextPollAt: session.status === 'STARTING' ? new Date(serverTime.getTime() + session.interval * 1000)
            : session.status === 'POLLING' ? session.pollLeaseUntil ?? session.nextPollAt : session.nextPollAt,
        status: session.status, errorCode: session.errorCode, errorMessage: session.errorMessage,
        authorizedEmail: session.authorizedEmail, createdBy: session.createdBy,
        createdAt: session.createdAt, completedAt: session.completedAt,
        serverTime,
    };
}

export function identityMatches(target: string, profile: unknown): { matches: boolean; authorizedEmail: string | null } {
    const record = profile && typeof profile === 'object' ? profile as Record<string, unknown> : {};
    const identities = [record.mail, record.userPrincipalName].filter(
        (value): value is string => typeof value === 'string' && value.length <= 255 && value.includes('@'),
    );
    const matched = identities.find((value) => value.trim().toLowerCase() === target.trim().toLowerCase());
    return {
        matches: matched !== undefined,
        authorizedEmail: matched ?? identities[0] ?? null,
    };
}

export function hasGraphMailReadWriteScope(scope: string | undefined): boolean {
    return typeof scope === 'string' && scope.split(/\s+/).some(
        (value) => /^(?:https:\/\/graph\.microsoft\.com\/)?Mail\.ReadWrite$/i.test(value),
    );
}

export function pollingOutcome(code: string, interval: number) {
    if (code === 'authorization_pending') return { status: 'PENDING' as const, interval };
    if (code === 'slow_down') return { status: 'PENDING' as const, interval: interval + 5 };
    if (code === 'temporarily_unavailable' || code === 'server_error') return { status: 'PENDING' as const, interval: Math.min(300, interval * 2) };
    if (code === 'authorization_declined' || code === 'access_denied') return { status: 'DECLINED' as const, interval };
    if (code === 'expired_token') return { status: 'EXPIRED' as const, interval };
    return { status: 'FAILED' as const, interval };
}

export function createReauthorizationService(db: PrismaClient = prisma, fetcher: typeof proxyFetch = proxyFetch) {
    async function cleanup() {
        const now = new Date();
        await db.emailReauthorization.updateMany({
            where: { status: { in: ACTIVE }, expiresAt: { lte: now } },
            data: { ...terminalData, status: 'EXPIRED', completedAt: now, errorCode: 'expired_token', errorMessage: '验证码已过期，请重试' },
        });
        // Never redeem a code again after an abandoned in-flight request: the
        // previous request might have consumed it. The administrator can restart.
        await db.emailReauthorization.updateMany({
            where: { status: 'POLLING', pollLeaseUntil: { lte: now } },
            data: { ...terminalData, status: 'FAILED', completedAt: now, errorCode: 'POLL_INTERRUPTED', errorMessage: '授权检查中断，请重新开始' },
        });
        await db.emailReauthorization.deleteMany({ where: { completedAt: { lt: new Date(now.getTime() - 30 * 86400000) } } });
    }

    async function find(id: string) {
        const session = await db.emailReauthorization.findUnique({ where: { id } });
        if (!session) throw new AppError('NOT_FOUND', '授权会话不存在', 404);
        return session;
    }

    async function finish(id: string, status: ReauthorizationStatus, code: string, message: string, claim?: string, authorizedEmail?: string | null) {
        await db.emailReauthorization.updateMany({
            where: { id, status: { in: ACTIVE }, ...(claim ? { pollClaim: claim } : {}) },
            data: { ...terminalData, status, errorCode: code, errorMessage: message, authorizedEmail, completedAt: new Date() },
        });
        return publicSession(await find(id));
    }

    return {
        async candidates() {
            await cleanup();
            const accounts = await db.emailAccount.findMany({
                where: { status: { not: 'DISABLED' }, errorMessage: { not: null } },
                select: { id: true, email: true, errorMessage: true,
                    group: { select: { name: true, fetchStrategy: true } },
                    reauthorizations: { where: { status: { in: ACTIVE } }, orderBy: { createdAt: 'desc' }, take: 1 },
                }, orderBy: { id: 'asc' },
            });
            return accounts.filter((account) => requiresReauthorization(account.errorMessage)).map((account) => {
                const supported = account.group?.fetchStrategy !== 'IMAP_ONLY';
                return {
                    emailId: account.id, email: account.email, groupName: account.group?.name ?? null, supported,
                    unsupportedReason: supported ? null : IMAP_ONLY_UNSUPPORTED_REASON,
                    // Existing errors may contain historical upstream bodies. Do not echo them.
                    reason: 'Microsoft 要求重新登录并授权',
                    activeSession: supported && account.reauthorizations[0] ? publicSession(account.reauthorizations[0]) : null,
                };
            });
        },

        async get(id: string) {
            await cleanup();
            return publicSession(await find(id));
        },

        async start(emailId: number, adminId: number) {
            await cleanup();
            const account = await db.emailAccount.findUnique({
                where: { id: emailId },
                select: {
                    id: true, email: true, clientId: true, tokenVersion: true,
                    status: true, errorMessage: true,
                    group: { select: { fetchStrategy: true } },
                },
            });
            if (!account || account.status === 'DISABLED') throw new AppError('ACCOUNT_UNAVAILABLE', '邮箱不存在或已禁用', 400);
            if (!requiresReauthorization(account.errorMessage)) throw new AppError('REAUTHORIZATION_NOT_REQUIRED', '该邮箱没有明确的重新授权错误，请刷新队列', 409);
            if (account.group?.fetchStrategy === 'IMAP_ONLY') {
                throw new AppError('REAUTHORIZATION_STRATEGY_UNSUPPORTED', IMAP_ONLY_UNSUPPORTED_REASON, 409);
            }
            const active = await db.emailReauthorization.findUnique({ where: { activeEmailId: emailId } });
            if (active) return publicSession(active);

            let session: EmailReauthorization;
            try {
                session = await db.emailReauthorization.create({ data: {
                    emailId, activeEmailId: emailId, email: account.email, clientId: account.clientId,
                    tokenVersion: account.tokenVersion, createdBy: adminId,
                    expiresAt: new Date(Date.now() + 60000), nextPollAt: new Date(Date.now() + 5000),
                } });
            } catch (error: unknown) {
                if (error && typeof error === 'object' && 'code' in error && error.code === 'P2002') {
                    const existing = await db.emailReauthorization.findUnique({ where: { activeEmailId: emailId } });
                    if (existing) return publicSession(existing);
                }
                throw new AppError('SESSION_START_FAILED', '无法创建授权会话，请重试', 503);
            }
            try {
                const response = await fetcher(`${MICROSOFT}/devicecode`, {
                    method: 'POST', signal: AbortSignal.timeout(15000),
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({ client_id: account.clientId, scope: DEVICE_SCOPES }).toString(),
                });
                if (!response.ok) {
                    const error = await readMicrosoftError(response);
                    return await finish(session.id, 'FAILED', error.code, error.message);
                }
                const parsed = deviceResponse.safeParse(await response.json());
                if (!parsed.success) return await finish(session.id, 'FAILED', 'INVALID_DEVICE_RESPONSE', 'Microsoft 返回了无效的设备码响应');
                const device = parsed.data;
                await db.emailReauthorization.updateMany({
                    where: { id: session.id, status: 'STARTING', expiresAt: { gt: new Date() } },
                    data: { status: 'PENDING', deviceCode: encrypt(device.device_code), userCode: device.user_code,
                        verificationUri: device.verification_uri, verificationUriComplete: device.verification_uri_complete ?? null,
                        expiresAt: new Date(Date.now() + device.expires_in * 1000), interval: device.interval,
                        nextPollAt: new Date(Date.now() + device.interval * 1000),
                    },
                });
                return publicSession(await find(session.id));
            } catch {
                return finish(session.id, 'FAILED', 'DEVICE_REQUEST_FAILED', '无法获取验证码，请检查网络或客户端的设备授权配置后重试');
            }
        },

        async cancel(id: string) {
            return finish(id, 'CANCELLED', 'CANCELLED', '已跳过此邮箱');
        },

        async poll(id: string) {
            await cleanup();
            const session = await find(id);
            const now = new Date();
            if (session.status !== 'PENDING' || session.nextPollAt > now) return publicSession(session);
            const claim = randomUUID();
            const claimed = await db.emailReauthorization.updateMany({
                where: { id, status: 'PENDING', nextPollAt: { lte: now }, expiresAt: { gt: now } },
                data: { status: 'POLLING', pollClaim: claim, pollLeaseUntil: new Date(now.getTime() + 60000) },
            });
            if (claimed.count !== 1) return publicSession(await find(id));

            const defer = async (interval: number, errorCode: string | null, errorMessage: string | null) => {
                await db.emailReauthorization.updateMany({
                    where: { id, status: 'POLLING', pollClaim: claim },
                    data: { status: 'PENDING', pollClaim: null, pollLeaseUntil: null, interval,
                        nextPollAt: new Date(Date.now() + interval * 1000), errorCode, errorMessage },
                });
                return publicSession(await find(id));
            };
            let tokens: z.infer<typeof tokenResponse>;
            try {
                if (!session.deviceCode) return await finish(id, 'FAILED', 'DEVICE_CODE_MISSING', '会话缺少验证码，请重试', claim);
                const response = await fetcher(`${MICROSOFT}/token`, {
                    method: 'POST', signal: AbortSignal.timeout(15000),
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({ client_id: session.clientId,
                        grant_type: 'urn:ietf:params:oauth:grant-type:device_code', device_code: decrypt(session.deviceCode) }).toString(),
                });
                if (!response.ok) {
                    const error = await readMicrosoftError(response);
                    const outcome = pollingOutcome(response.status === 429 ? 'slow_down' : response.status >= 500 ? 'server_error' : error.code, session.interval);
                    if (outcome.status === 'PENDING') return await defer(outcome.interval,
                        error.code === 'authorization_pending' ? null : error.code,
                        error.code === 'authorization_pending' ? null : error.message);
                    return await finish(id, outcome.status, error.code, error.message, claim);
                }
                const parsed = tokenResponse.safeParse(await response.json());
                if (!parsed.success) return await finish(id, 'FAILED', 'TOKEN_RESPONSE_INCOMPLETE', '授权响应缺少必需的令牌，请重新授权', claim);
                tokens = parsed.data;
                if (!hasGraphMailReadWriteScope(tokens.scope)) {
                    return await finish(id, 'FAILED', 'GRAPH_MAIL_READWRITE_SCOPE_MISSING', 'Microsoft 未授予 Graph Mail.ReadWrite 权限，未保存令牌，请重新授权', claim);
                }
            } catch {
                return defer(Math.min(300, Math.max(5, session.interval * 2)), 'NETWORK_ERROR', '网络暂时不可用，稍后自动重试');
            }

            // After redemption the device code may be single-use. If /me fails,
            // discard the tokens and require a new session, never redeem twice.
            try {
                const response = await fetcher('https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName', {
                    headers: { Authorization: `Bearer ${tokens.access_token}` }, signal: AbortSignal.timeout(15000),
                });
                if (!response.ok) return await finish(id, 'FAILED', 'IDENTITY_CHECK_FAILED', '无法核验 Microsoft 账号，请检查 User.Read 权限后重试', claim);
                const identity = identityMatches(session.email, await response.json());
                if (!identity.matches) return await finish(id, 'MISMATCH', 'ACCOUNT_MISMATCH', '登录的 Microsoft 账号与当前邮箱不一致，未保存令牌。请切换账号后重试', claim, identity.authorizedEmail);

                const encrypted = encrypt(tokens.refresh_token);
                await db.$transaction(async (tx) => {
                    // Lock/claim the session first so cancellation and expiry are
                    // ordered atomically against the credential write.
                    const won = await tx.emailReauthorization.updateMany({
                        where: { id, status: 'POLLING', pollClaim: claim, expiresAt: { gt: new Date() } },
                        data: { ...terminalData, status: 'SUCCEEDED', authorizedEmail: identity.authorizedEmail,
                            completedAt: new Date(), errorCode: null, errorMessage: null },
                    });
                    if (won.count !== 1) return;
                    const saved = await tx.emailAccount.updateMany({
                        where: { id: session.emailId, tokenVersion: session.tokenVersion,
                            email: session.email, clientId: session.clientId, status: { not: 'DISABLED' },
                            OR: [
                                { groupId: null },
                                { group: { is: { fetchStrategy: { not: 'IMAP_ONLY' } } } },
                            ],
                        },
                        data: { refreshToken: encrypted, tokenRefreshedAt: new Date(),
                            status: 'ACTIVE', errorMessage: null, tokenVersion: { increment: 1 } },
                    });
                    if (saved.count !== 1) await tx.emailReauthorization.update({ where: { id },
                        data: { status: 'FAILED', errorCode: 'ACCOUNT_CHANGED', errorMessage: '账号数据已变化，未覆盖新的令牌，请重试' },
                    });
                });
                await cleanup();
                return publicSession(await find(id));
            } catch {
                return finish(id, 'FAILED', 'AUTHORIZATION_SAVE_FAILED', '账号核验或保存失败，请重新开始授权', claim);
            }
        },
    };
}

export const reauthorizationService = createReauthorizationService();
