import { createHash, randomBytes } from 'node:crypto';
import type { PrismaClient } from '@prisma/client';
import { z } from 'zod';
import prisma from '../../lib/prisma.js';
import { getRedis } from '../../lib/redis.js';
import { decrypt } from '../../lib/crypto.js';
import { requiresReauthorization } from '../../lib/microsoft-oauth.js';
import { AppError } from '../../plugins/error.js';

export const HELPER_ORIGIN = 'https://outlook.wujiaqiao.dpdns.org';
export const helperSecret = z.string().regex(/^[A-Za-z0-9_-]{43}$/);
const bindingSchema = z.object({
    sessionId: z.string().uuid(), emailId: z.number().int().positive(), adminId: z.number().int().positive(),
    tokenVersion: z.number().int().nonnegative(), ua: z.string().regex(/^[a-f0-9]{64}$/),
    expiresAt: z.number().int().positive(),
}).strict();
type Binding = z.infer<typeof bindingSchema>;
type Audit = (event: { action: string; adminId: number; emailId: number; sessionId: string }) => void;
export interface HelperStore {
    get(key: string): Promise<string | null>;
    take(key: string): Promise<string | null>;
    put(key: string, value: string, ttl: number): Promise<boolean>;
    issue(adminKey: string, expected: string | null, sessionKey: string, ticketKey: string,
        value: string, ttl: number, ticketTtl: number): Promise<boolean>;
}
const prefix = 'reauth-helper:v1:';
export const helperDigest = (value: string) => createHash('sha256').update(value).digest('hex');
const invalid = () => new AppError('HELPER_INVALID', '助手凭据无效或已过期，请回后台人工检查当前会话', 401);
const unavailable = () => new AppError('HELPER_UNAVAILABLE', '安全存储暂时不可用，助手已停止', 503);

// No process-memory fallback, secret-bearing cache helpers or raw Redis errors.
function redisStore(): HelperStore {
    const client = getRedis();
    if (!client) throw unavailable();
    return {
        get: (key) => client.get(key),
        // Redis >= 6.2. GETDEL gives exactly one winner across all instances.
        take: (key) => client.getdel(key),
        put: async (key, value, ttl) => await client.set(key, value, 'PX', ttl, 'NX') === 'OK',
        issue: async (adminKey, expected, sessionKey, ticketKey, value, ttl, ticketTtl) =>
            await client.eval(`
                local previous = redis.call('GET', KEYS[1]) or ''
                if previous ~= ARGV[1] or redis.call('EXISTS', KEYS[2]) == 1 then return 0 end
                redis.call('SET', KEYS[1], ARGV[2], 'PX', ARGV[3])
                redis.call('SET', KEYS[2], '1', 'PX', ARGV[3])
                redis.call('SET', KEYS[3], ARGV[2], 'PX', ARGV[4])
                return 1`, 3, adminKey, sessionKey, ticketKey, expected ?? '', value, ttl, ticketTtl) === 1,
    };
}

export function createHelperService(db: PrismaClient = prisma, getStore: () => HelperStore = redisStore,
    now: () => number = Date.now) {
    async function stored<T>(operation: (store: HelperStore) => Promise<T>): Promise<T> {
        try { return await operation(getStore()); } catch { throw unavailable(); }
    }
    function read(raw: string | null, ua: string): Binding {
        try {
            if (!ua || ua.length > 1024) throw invalid();
            const value = bindingSchema.parse(JSON.parse(raw ?? 'null'));
            if (value.expiresAt <= now() || value.ua !== helperDigest(ua)) throw invalid();
            return value;
        } catch { throw invalid(); }
    }
    async function validationState(binding: Binding) {
        const session = await db.emailReauthorization.findUnique({ where: { id: binding.sessionId } });
        const [account, admin] = await Promise.all([
            db.emailAccount.findUnique({ where: { id: binding.emailId }, select: {
                id: true, email: true, clientId: true, tokenVersion: true, status: true, errorMessage: true,
                group: { select: { fetchStrategy: true } },
            } }),
            db.admin.findUnique({ where: { id: binding.adminId }, select: { role: true, status: true } }),
        ]);
        return { session, account, admin };
    }
    async function validate(binding: Binding, allowSuccess = false) {
        let state = await validationState(binding);
        const accepts = (value: typeof state) => {
            const { session, account, admin } = value;
            const succeeded = allowSuccess && session?.status === 'SUCCEEDED';
            return !!session && !!account && admin?.role === 'SUPER_ADMIN' && admin.status === 'ACTIVE' &&
                binding.expiresAt > now() && session.expiresAt.getTime() > now() &&
                (succeeded || ['PENDING', 'POLLING'].includes(session.status)) &&
                (session.status !== 'POLLING' || !!session.pollLeaseUntil && session.pollLeaseUntil.getTime() > now()) &&
                session.createdBy === binding.adminId && session.emailId === binding.emailId &&
                session.activeEmailId === (succeeded ? null : binding.emailId) && session.tokenVersion === binding.tokenVersion &&
                account.tokenVersion === binding.tokenVersion + (succeeded ? 1 : 0) && account.email === session.email &&
                account.clientId === session.clientId && account.status !== 'DISABLED' &&
                account.group?.fetchStrategy !== 'IMAP_ONLY' &&
                (succeeded || requiresReauthorization(account.errorMessage) && !!session.userCode);
        };
        if (accepts(state)) return state.session!;
        // A successful reauthorization commits the account version and terminal session together. Under
        // READ COMMITTED, the session query can see the old PENDING row while the later account query
        // sees that commit. Once the increment is visible, a complete reread must see the same commit.
        const successCommitMayHaveInterleaved = allowSuccess &&
            !!state.session && ['PENDING', 'POLLING'].includes(state.session.status) &&
            state.session.activeEmailId === binding.emailId && state.session.tokenVersion === binding.tokenVersion &&
            state.account?.tokenVersion === binding.tokenVersion + 1;
        if (successCommitMayHaveInterleaved) {
            state = await validationState(binding);
            if (accepts(state) && state.session?.status === 'SUCCEEDED') return state.session;
        }
        throw invalid();
    }
    async function capability(secret: string, ua: string, allowSuccess = false) {
        if (!helperSecret.safeParse(secret).success) throw invalid();
        const binding = read(await stored((store) => store.get(`${prefix}cap:${helperDigest(secret)}`)), ua);
        const session = await validate(binding, allowSuccess);
        return { binding, session };
    }
    function audit(action: string, binding: Binding, log: Audit) {
        log({ action, adminId: binding.adminId, emailId: binding.emailId, sessionId: binding.sessionId });
    }
    return {
        async issue(sessionId: string, adminId: number, ua: string, log: Audit) {
            if (!ua || ua.length > 1024) throw invalid();
            const session = await db.emailReauthorization.findUnique({ where: { id: sessionId } });
            if (!session) throw invalid();
            const binding: Binding = { sessionId, emailId: session.emailId, adminId,
                tokenVersion: session.tokenVersion, ua: helperDigest(ua), expiresAt: Math.min(now() + 90000, session.expiresAt.getTime()) };
            await validate(binding);
            const adminKey = `${prefix}admin:${adminId}`;
            const previous = await stored((store) => store.get(adminKey));
            if (previous) {
                let previousId: string;
                try { previousId = bindingSchema.parse(JSON.parse(previous)).sessionId; } catch { throw invalid(); }
                const prior = await db.emailReauthorization.findUnique({ where: { id: previousId } });
                if (prior && (['STARTING', 'PENDING', 'POLLING'].includes(prior.status) && prior.expiresAt.getTime() > now() ||
                    prior.status === 'SUCCEEDED' && (!prior.completedAt || prior.completedAt.getTime() + 10000 > now()))) {
                    throw new AppError('HELPER_BUSY', '请先完成当前会话；成功后至少等待 10 秒', 409);
                }
            }
            const ticket = randomBytes(32).toString('base64url');
            const accepted = await stored((store) => store.issue(adminKey, previous, `${prefix}issued:${sessionId}`,
                `${prefix}ticket:${helperDigest(ticket)}`, JSON.stringify(binding),
                Math.max(1, session.expiresAt.getTime() - now() + 10000), Math.max(1, binding.expiresAt - now())));
            if (!accepted) throw new AppError('HELPER_ALREADY_ISSUED', '此会话已签发助手凭据，请人工继续，或取消后重新创建会话', 409);
            audit('email_reauthorization.helper_ticket_issued', binding, log);
            return { ticket, expiresAt: new Date(binding.expiresAt).toISOString() };
        },
        async claim(ticket: string, ua: string) {
            if (!helperSecret.safeParse(ticket).success) throw invalid();
            const binding = read(await stored((store) => store.take(`${prefix}ticket:${helperDigest(ticket)}`)), ua);
            const session = await validate(binding);
            binding.expiresAt = Math.min(now() + 300000, session.expiresAt.getTime());
            const secret = randomBytes(32).toString('base64url');
            const digest = helperDigest(secret);
            const value = JSON.stringify(binding);
            const ttl = Math.max(1, binding.expiresAt - now());
            // A partial Redis failure only leaves an unexposed random capability.
            if (!await stored((store) => store.put(`${prefix}cap:${digest}`, value, ttl)) ||
                !await stored((store) => store.put(`${prefix}password:${digest}`, value, ttl))) throw unavailable();
            return { capability: secret, sessionId: binding.sessionId, expiresAt: new Date(binding.expiresAt).toISOString() };
        },
        async current(secret: string, ua: string) {
            const { binding, session } = await capability(secret, ua, true);
            return { sessionId: session.id, email: session.email, userCode: session.status === 'SUCCEEDED' ? null : session.userCode!,
                clientId: session.clientId, status: session.status, expiresAt: new Date(binding.expiresAt).toISOString() };
        },
        async password(secret: string, ua: string, log: Audit) {
            const { binding } = await capability(secret, ua);
            const consumed = read(await stored((store) => store.take(`${prefix}password:${helperDigest(secret)}`)), ua);
            if (JSON.stringify(consumed) !== JSON.stringify(binding)) throw invalid();
            await validate(binding);
            // Fetch ciphertext only after the atomic one-shot permission is won.
            const account = await db.emailAccount.findUnique({ where: { id: binding.emailId }, select: { password: true } });
            await validate(binding);
            let password: string;
            try {
                if (!account?.password) throw invalid();
                password = decrypt(account.password);
                if (!password) throw invalid();
            } catch { throw new AppError('HELPER_PASSWORD_UNAVAILABLE', '保存的密码不可用，请人工处理', 409); }
            audit('email_reauthorization.helper_password_delivered', binding, log);
            return { password };
        },
    };
}

export const helperService = createHelperService();
