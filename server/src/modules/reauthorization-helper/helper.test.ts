import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
import Fastify from 'fastify';
import pino from 'pino';
import type { EmailReauthorization, PrismaClient } from '@prisma/client';
import type { HelperStore } from './helper.service.js';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.JWT_SECRET = 'test-jwt-secret-for-helper-00000000';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
const { createHelperService, helperDigest, helperService } = await import('./helper.service.js');
const { encrypt } = await import('../../lib/crypto.js');
const { default: helperRoutes } = await import('./helper.routes.js');
const { default: adminRoutes } = await import('../email-reauthorization/reauthorization.routes.js');
const { AppError } = await import('../../plugins/error.js');
const ID = '12345678-1234-4123-8123-123456789012';
const UA = 'test-browser-agent';
const CLEAR = 'test-only-mailbox-credential';

async function runMicrosoftScenario(options: {
    url: string;
    tabTask?: Record<string, unknown>;
    values?: Record<string, unknown>;
    view?: 'empty' | 'device' | 'email' | 'password' | 'consent' | 'password-consent' | 'password-continue';
    autoTicket?: boolean;
    stopAfterClicks?: number;
    formAction?: string;
    formMethod?: string;
    submitFormAction?: string;
    submitFormMethod?: string;
    currentClientId?: string;
    saveMutations?: Record<number, { url?: string; formAction?: string; formMethod?: string;
        submitFormAction?: string; submitFormMethod?: string }>;
}) {
    const source = await readFile(new URL('../../../../web/public/gongxi-mail-reauthorization.user.js', import.meta.url), 'utf8');
    const executable = source.replace(
        /[ ]{4}if \(location\.origin === ORIGIN\) void adminMain\(\);\r?\n[ ]{4}else if \(MS_HOSTS\.includes\(location\.hostname\)\) void microsoftMain\(\);/,
        '    module.exports.microsoftMain = microsoftMain;',
    );
    assert.notEqual(executable, source, 'userscript test hook must replace the runtime dispatch');
    const parsed = new URL(options.url);
    const locationValue = { origin: parsed.origin, protocol: parsed.protocol, hostname: parsed.hostname,
        pathname: parsed.pathname, hash: parsed.hash, search: parsed.search, href: parsed.href };
    const setLocation = (raw: string) => {
        const next = new URL(raw, locationValue.href);
        Object.assign(locationValue, { origin: next.origin, protocol: next.protocol, hostname: next.hostname,
            pathname: next.pathname, hash: next.hash, search: next.search, href: next.href });
    };
    class TestInput {
        disabled = false;
        innerText = '';
        textContent = '';
        form: unknown;
        private storedValue = '';
        get value() { return this.storedValue; }
        set value(value: string) { this.storedValue = value; }
        getClientRects() { return [{}]; }
        getAttribute() { return null; }
        closest() { return this.form; }
        dispatchEvent() { return true; }
    }
    const makeNode = (text = '') => ({
        disabled: false, innerText: text, textContent: text, value: '',
        getClientRects: () => [{}], getAttribute: () => null,
    });
    const identity = makeNode('target@outlook.com');
    const email = new TestInput();
    const password = new TestInput();
    const device = new TestInput();
    const submitText = ['consent', 'password-consent'].includes(options.view ?? '') ? 'Accept' :
        options.view === 'password-continue' || options.view === 'device' ? 'Continue' :
            options.view === 'email' ? 'Next' : 'Sign in';
    const submit = makeNode(submitText);
    const submitAttributes: Record<string, string> = {};
    if (options.submitFormAction !== undefined) submitAttributes.formaction = options.submitFormAction;
    if (options.submitFormMethod !== undefined) submitAttributes.formmethod = options.submitFormMethod;
    const submitListeners: Array<{ listener: () => void; once: boolean }> = [];
    let submitEvents = 0;
    const form = {
        action: options.formAction ?? parsed.href,
        method: options.formMethod ?? 'post',
        addEventListener: (type: string, listener: () => void, settings?: { once?: boolean }) => {
            if (type === 'submit') submitListeners.push({ listener, once: !!settings?.once });
        },
    };
    email.form = form;
    password.form = form;
    device.form = form;
    let clicks = 0;
    let clickReached: (() => void) | undefined;
    const reached = new Promise<void>((resolve) => { clickReached = resolve; });
    Object.assign(submit, {
        form,
        formAction: options.submitFormAction ?? '',
        formMethod: options.submitFormMethod ?? '',
        getAttribute: (name: string) => submitAttributes[name] ?? null,
        closest: () => form,
        click: () => {
        clicks++;
        if (options.view === 'device') {
            submitEvents++;
            for (const registered of [...submitListeners]) registered.listener();
            for (let index = submitListeners.length - 1; index >= 0; index--) {
                if (submitListeners[index].once) submitListeners.splice(index, 1);
            }
        }
        if (clicks >= (options.stopAfterClicks ?? Number.POSITIVE_INFINITY)) clickReached?.();
        },
    });
    const bodyText = ['consent', 'password-consent'].includes(options.view ?? '') ? 'Permissions requested by this app' :
        options.view === 'password-continue' ? 'Continue signing in to this app' : '';
    const querySelectorAll = (selector: string) => {
        if (selector.startsWith('#displayName,')) return ['password', 'consent', 'password-consent', 'password-continue'].includes(options.view ?? '') ? [identity] : [];
        if (selector === '#i0116[name="loginfmt"]') return options.view === 'email' ? [email] : [];
        if (selector === 'input#i0118[name="passwd"][type="password"]') return ['password', 'password-consent', 'password-continue'].includes(options.view ?? '') ? [password] : [];
        if (selector === '#otc, input[name="user_code"]') return options.view === 'device' ? [device] : [];
        if (selector === '#idSIButton9, #idSubmit_Consent, #idBtn_Accept') return options.view === 'empty' ? [] : [submit];
        return [];
    };
    const makeElement = () => ({
        id: '', style: { cssText: '' }, textContent: '',
        append: () => undefined, addEventListener: () => undefined,
        attachShadow: () => ({ append: () => undefined }),
    });
    const documentValue = {
        body: { innerText: bodyText, append: () => undefined },
        createElement: makeElement,
        querySelectorAll,
    };
    const windowValue: Record<string, unknown> = { addEventListener: () => undefined };
    windowValue.top = windowValue;
    windowValue.self = windowValue;
    const tabState = options.tabTask ? { gongxiHelper: { ...options.tabTask } } : {};
    const values = new Map(Object.entries(options.values ?? {}));
    const requests: string[] = [];
    const handoffEvents: string[] = [];
    const moduleValue: { exports: { microsoftMain?: () => Promise<void> } } = { exports: {} };
    const sandbox = {
        module: moduleValue,
        document: documentValue,
        window: windowValue,
        location: locationValue,
        history: { replaceState: (_state: unknown, _unused: string, url: string) => { setLocation(url); } },
        URL,
        Event: class { constructor(public type: string, public options?: object) {} },
        HTMLInputElement: TestInput,
        setTimeout: (callback: (...args: unknown[]) => void, delay?: number, ...args: unknown[]) => {
            if (clicks >= (options.stopAfterClicks ?? Number.POSITIVE_INFINITY)) return 0;
            return setTimeout(callback, delay, ...args);
        },
        clearTimeout,
        Date,
        GM_getTab: (done: (value: object) => void) => done(tabState),
        GM_saveTab: (() => {
            let saves = 0;
            return (_value: object, done?: () => void) => {
                saves++;
                const mutation = options.saveMutations?.[saves];
                if (mutation?.url) setLocation(mutation.url);
                if (mutation?.formAction !== undefined) form.action = mutation.formAction;
                if (mutation?.formMethod !== undefined) form.method = mutation.formMethod;
                if (mutation?.submitFormAction !== undefined) {
                    submitAttributes.formaction = mutation.submitFormAction;
                    Object.assign(submit, { formAction: mutation.submitFormAction });
                }
                if (mutation?.submitFormMethod !== undefined) {
                    submitAttributes.formmethod = mutation.submitFormMethod;
                    Object.assign(submit, { formMethod: mutation.submitFormMethod });
                }
                done?.();
            };
        })(),
        GM_getValue: (key: string, fallback: unknown) => {
            if (values.has(key)) return values.get(key);
            if (options.autoTicket && key.startsWith('gongxi-helper-ticket-v2:')) {
                const binding = values.get(key.replace('gongxi-helper-ticket-v2:', 'gongxi-helper-binding-v2:')) as
                    { runId?: unknown; sessionId?: unknown; bindingId?: unknown; boundAt?: unknown } | undefined;
                if (binding) {
                    handoffEvents.push('ticket-published');
                    return { ...binding, ticket: 't'.repeat(43), publishedAt: binding.boundAt, expiresAt: Date.now() + 60000 };
                }
            }
            return fallback;
        },
        GM_setValue: (key: string, value: unknown) => {
            values.set(key, value);
            if (key.startsWith('gongxi-helper-binding-v2:')) handoffEvents.push('binding-saved');
        },
        GM_deleteValue: (key: string) => { values.delete(key); },
        GM_xmlhttpRequest: (request: {
            url: string;
            onload: (response: { finalUrl: string; status: number; response: object; responseText: string }) => void;
            onerror: () => void;
        }) => {
            const path = request.url.split('/').at(-1)!;
            requests.push(path);
            if (path === 'claim') handoffEvents.push('ticket-claimed');
            if (path === 'claim') {
                request.onload({ finalUrl: request.url, status: 200, response: { success: true, data: {
                    sessionId: ID, capability: 'c'.repeat(43), expiresAt: new Date(Date.now() + 60000).toISOString(),
                } }, responseText: '' });
            } else if (path === 'current') {
                request.onload({ finalUrl: request.url, status: 200, response: { success: true, data: {
                    sessionId: ID, email: 'target@outlook.com', userCode: 'ABCD-EFGH', status: 'PENDING',
                    clientId: options.currentClientId ?? 'client',
                    expiresAt: new Date(Date.now() + 60000).toISOString(),
                } }, responseText: '' });
            } else if (path === 'password') {
                request.onload({ finalUrl: request.url, status: 200, response: { success: true, data: {
                    password: CLEAR,
                } }, responseText: '' });
            } else {
                request.onerror();
            }
        },
    };
    runInNewContext(executable, sandbox, { timeout: 1000 });
    assert.ok(moduleValue.exports.microsoftMain);
    const execution = moduleValue.exports.microsoftMain();
    if (options.stopAfterClicks) await Promise.race([execution, reached]);
    else await execution;
    return { requests, clicks, submitEvents, inputValues: { device: device.value, email: email.value, password: password.value },
        handoffEvents, tabState, values };
}

function fixture() {
    let time = Date.now();
    const session = { id: ID, emailId: 7, activeEmailId: 7, email: 'target@outlook.com', clientId: 'client',
        tokenVersion: 4, createdBy: 1, status: 'PENDING', userCode: 'ABCD-EFGH', deviceCode: encrypt('secret-device'),
        expiresAt: new Date(time + 900000), pollLeaseUntil: null, completedAt: null,
    } as EmailReauthorization;
    const account = { id: 7, email: session.email, clientId: 'client', tokenVersion: 4, status: 'ERROR',
        errorMessage: 'AADSTS70000', password: encrypt(CLEAR), group: { fetchStrategy: 'GRAPH_FIRST' } };
    const admin = { role: 'SUPER_ADMIN', status: 'ACTIVE' };
    let passwordReads = 0;
    const db = {
        emailReauthorization: { findUnique: async ({ where }: { where: { id: string } }) => where.id === session.id ? { ...session } : null },
        emailAccount: { findUnique: async ({ select }: { select: { password?: boolean } }) => {
            if (select.password) passwordReads++;
            return { ...account };
        } },
        admin: { findUnique: async () => ({ ...admin }) },
    } as unknown as PrismaClient;
    const entries = new Map<string, { value: string; expires: number }>();
    function get(key: string) {
        const value = entries.get(key);
        return value && value.expires > time ? value.value : null;
    }
    const store: HelperStore = {
        get: async (key) => get(key),
        take: async (key) => { const value = get(key); entries.delete(key); return value; },
        put: async (key, value, ttl) => {
            if (get(key)) return false;
            entries.set(key, { value, expires: time + ttl }); return true;
        },
        issue: async (adminKey, expected, sessionKey, ticketKey, value, ttl, ticketTtl) => {
            if (get(adminKey) !== expected || get(sessionKey)) return false;
            entries.set(adminKey, { value, expires: time + ttl });
            entries.set(sessionKey, { value: '1', expires: time + ttl });
            entries.set(ticketKey, { value, expires: time + ticketTtl });
            return true;
        },
    };
    const events: Record<string, unknown>[] = [];
    const log = (value: Record<string, unknown>) => { events.push(value); };
    const service = createHelperService(db, () => store, () => time);
    return { service, db, store, session, account, admin, entries, events, log,
        now: () => time, advance: (ms: number) => { time += ms; }, passwordReads: () => passwordReads };
}

void test('ticket binds session/account/admin/version/UA, stores digests only, and claims once', async () => {
    const f = fixture();
    const issued = await f.service.issue(ID, 1, UA, f.log);
    assert.equal(Date.parse(issued.expiresAt) - f.now(), 90000);
    const serialized = JSON.stringify([...f.entries]);
    assert.equal(serialized.includes(issued.ticket), false);
    assert.equal(serialized.includes(UA), false);
    assert.doesNotMatch(serialized, /secret-device|ABCD-EFGH|target@outlook/);
    const ticketEntry = [...f.entries].find(([key]) => key.includes(':ticket:'))!;
    assert.ok(ticketEntry[0].endsWith(helperDigest(issued.ticket)));
    assert.deepEqual(Object.keys(JSON.parse(ticketEntry[1].value) as object).sort(), ['adminId', 'emailId', 'expiresAt', 'sessionId', 'tokenVersion', 'ua']);
    const results = await Promise.allSettled([f.service.claim(issued.ticket, UA), f.service.claim(issued.ticket, UA)]);
    assert.equal(results.filter((result) => result.status === 'fulfilled').length, 1);
    const claimed = results.find((result) => result.status === 'fulfilled')!;
    assert.equal(claimed.status, 'fulfilled');
    if (claimed.status !== 'fulfilled') return;
    assert.equal(Date.parse(claimed.value.expiresAt) - f.now(), 300000);
    assert.equal(JSON.stringify([...f.entries]).includes(claimed.value.capability), false);
    const current = await f.service.current(claimed.value.capability, UA);
    assert.deepEqual(Object.keys(current).sort(), ['clientId', 'email', 'expiresAt', 'sessionId', 'status', 'userCode']);
    assert.equal(current.clientId, f.session.clientId);
    await assert.rejects(f.service.current(issued.ticket, UA), { code: 'HELPER_INVALID' });
    await assert.rejects(f.service.claim(claimed.value.capability, UA), { code: 'HELPER_INVALID' });
    await assert.rejects(f.service.issue(ID, 1, UA, f.log));
});

void test('password has one atomic winner across concurrent requests and audits only identifiers', async () => {
    const f = fixture();
    const { ticket } = await f.service.issue(ID, 1, UA, f.log);
    const { capability } = await f.service.claim(ticket, UA);
    const results = await Promise.allSettled(Array.from({ length: 5 }, () => f.service.password(capability, UA, f.log)));
    assert.equal(results.filter((result) => result.status === 'fulfilled').length, 1);
    const result = results.find((value) => value.status === 'fulfilled');
    assert.ok(result?.status === 'fulfilled');
    assert.deepEqual(result.value, { password: CLEAR });
    assert.equal(f.passwordReads(), 1);
    assert.equal(f.events.length, 2);
    for (const event of f.events) assert.deepEqual(Object.keys(event).sort(), ['action', 'adminId', 'emailId', 'sessionId']);
    for (const secret of [CLEAR, ticket, capability, 'ABCD-EFGH', 'secret-device']) {
        assert.equal(JSON.stringify(f.events).includes(secret), false);
        assert.equal(JSON.stringify([...f.entries]).includes(secret), false);
    }
});

for (const mutation of ['emailId', 'createdBy', 'tokenVersion', 'accountVersion', 'email', 'clientId', 'disabled', 'adminDisabled', 'role', 'cancelled', 'imap', 'expired', 'lease'] as const) {
    void test(`each capability request rechecks ${mutation}`, async () => {
        const f = fixture();
        const { ticket } = await f.service.issue(ID, 1, UA, f.log);
        const { capability } = await f.service.claim(ticket, UA);
        if (mutation === 'emailId') f.session.emailId++;
        if (mutation === 'createdBy') f.session.createdBy = 2;
        if (mutation === 'tokenVersion') f.session.tokenVersion++;
        if (mutation === 'accountVersion') f.account.tokenVersion++;
        if (mutation === 'email') f.account.email = 'other@outlook.com';
        if (mutation === 'clientId') f.account.clientId = 'another';
        if (mutation === 'disabled') f.account.status = 'DISABLED';
        if (mutation === 'adminDisabled') f.admin.status = 'DISABLED';
        if (mutation === 'role') f.admin.role = 'ADMIN';
        if (mutation === 'cancelled') f.session.status = 'CANCELLED';
        if (mutation === 'imap') f.account.group.fetchStrategy = 'IMAP_ONLY';
        if (mutation === 'expired') f.session.expiresAt = new Date(f.now());
        if (mutation === 'lease') { f.session.status = 'POLLING'; f.session.pollLeaseUntil = new Date(f.now()); }
        await assert.rejects(f.service.current(capability, UA), { code: 'HELPER_INVALID' });
        await assert.rejects(f.service.password(capability, UA, f.log), { code: 'HELPER_INVALID' });
        assert.equal(f.passwordReads(), 0);
    });
}

void test('tampered binding, wrong user agent and malformed credentials fail closed', async () => {
    const f = fixture();
    await assert.rejects(f.service.issue(ID, 2, UA, f.log), { code: 'HELPER_INVALID' });
    const { ticket } = await f.service.issue(ID, 1, UA, f.log);
    await assert.rejects(f.service.claim(ticket, 'other-browser'), { code: 'HELPER_INVALID' });
    await assert.rejects(f.service.claim(ticket, UA), { code: 'HELPER_INVALID' });
    await assert.rejects(f.service.claim('raw-invalid-secret', UA), { code: 'HELPER_INVALID' });
    const g = fixture();
    const issued = await g.service.issue(ID, 1, UA, g.log);
    const { capability } = await g.service.claim(issued.ticket, UA);
    await assert.rejects(g.service.current(capability, 'other-browser'), { code: 'HELPER_INVALID' });
    const record = [...g.entries].find(([key]) => key.includes(':cap:'))![1];
    record.value = record.value.replace('"emailId":7', '"emailId":8');
    await assert.rejects(g.service.current(capability, UA), { code: 'HELPER_INVALID' });
});

void test('ticket/capability expiration and device expiry bound credentials', async () => {
    const f = fixture();
    const { ticket } = await f.service.issue(ID, 1, UA, f.log);
    f.advance(90000);
    await assert.rejects(f.service.claim(ticket, UA), { code: 'HELPER_INVALID' });
    const g = fixture();
    g.session.expiresAt = new Date(g.now() + 40000);
    const short = await g.service.issue(ID, 1, UA, g.log);
    assert.equal(Date.parse(short.expiresAt), g.session.expiresAt.getTime());
    const cap = await g.service.claim(short.ticket, UA);
    assert.equal(Date.parse(cap.expiresAt), g.session.expiresAt.getTime());
    g.advance(40000);
    await assert.rejects(g.service.password(cap.capability, UA, g.log), { code: 'HELPER_INVALID' });
    const h = fixture();
    const next = await h.service.claim((await h.service.issue(ID, 1, UA, h.log)).ticket, UA);
    h.advance(300000);
    await assert.rejects(h.service.current(next.capability, UA), { code: 'HELPER_INVALID' });
});

void test('success can be observed but never releases a password or reissues the session', async () => {
    const f = fixture();
    const { ticket } = await f.service.issue(ID, 1, UA, f.log);
    const { capability } = await f.service.claim(ticket, UA);
    f.session.status = 'SUCCEEDED'; f.session.activeEmailId = null; f.session.userCode = null;
    f.session.completedAt = new Date(f.now()); f.account.tokenVersion++;
    assert.equal((await f.service.current(capability, UA)).status, 'SUCCEEDED');
    assert.equal((await f.service.current(capability, UA)).userCode, null);
    await assert.rejects(f.service.password(capability, UA, f.log), { code: 'HELPER_INVALID' });
    await assert.rejects(f.service.issue(ID, 1, UA, f.log));
});

void test('current recovers from a committed success interleaved between session and account reads', async (t) => {
    const f = fixture();
    const { capability } = await f.service.claim((await f.service.issue(ID, 1, UA, f.log)).ticket, UA);
    const staleSession = { ...f.session };
    f.session.status = 'SUCCEEDED';
    f.session.activeEmailId = null;
    f.session.userCode = null;
    f.session.completedAt = new Date(f.now());
    f.account.tokenVersion++;
    let sessionReads = 0;
    t.mock.method(f.db.emailReauthorization, 'findUnique', async () => {
        sessionReads++;
        return sessionReads === 1 ? { ...staleSession } : { ...f.session };
    });
    const current = await f.service.current(capability, UA);
    assert.equal(current.status, 'SUCCEEDED');
    assert.equal(current.userCode, null);
    assert.ok(sessionReads >= 2, 'an inconsistent transition must trigger a complete reread');
    await assert.rejects(f.service.password(capability, UA, f.log), { code: 'HELPER_INVALID' });
    await assert.rejects(f.service.password(capability, UA, f.log), { code: 'HELPER_INVALID' });
    assert.equal(f.passwordReads(), 0);
});

void test('Redis failures have no fallback and a lost password response cannot be retried', async () => {
    const f = fixture();
    const service = createHelperService(f.db, () => { throw new Error('redis-raw-secret'); });
    await assert.rejects(service.issue(ID, 1, UA, f.log), { code: 'HELPER_UNAVAILABLE', message: '安全存储暂时不可用，助手已停止' });
    const { ticket } = await f.service.issue(ID, 1, UA, f.log);
    const { capability } = await f.service.claim(ticket, UA);
    f.account.password = 'invalid-ciphertext';
    await assert.rejects(f.service.password(capability, UA, f.log), { code: 'HELPER_PASSWORD_UNAVAILABLE' });
    f.account.password = encrypt(CLEAR);
    await assert.rejects(f.service.password(capability, UA, f.log), { code: 'HELPER_INVALID' });
    assert.equal(f.passwordReads(), 1);
});

void test('admin guard serializes issuance and enforces ten seconds after confirmed success', async (t) => {
    const f = fixture();
    await f.service.issue(ID, 1, UA, f.log);
    const prior = { ...f.session };
    f.session.id = '22345678-1234-4123-8123-123456789012';
    t.mock.method(f.db.emailReauthorization, 'findUnique', async ({ where }: { where: { id: string } }) =>
        where.id === ID ? { ...prior } : { ...f.session });
    await assert.rejects(f.service.issue(f.session.id, 1, UA, f.log), { code: 'HELPER_BUSY' });
    prior.status = 'SUCCEEDED'; prior.completedAt = new Date(f.now());
    f.session.tokenVersion++; f.account.tokenVersion++;
    f.advance(9999);
    await assert.rejects(f.service.issue(f.session.id, 1, UA, f.log), { code: 'HELPER_BUSY' });
    f.advance(1);
    const results = await Promise.allSettled([f.service.issue(f.session.id, 1, UA, f.log), f.service.issue(f.session.id, 1, UA, f.log)]);
    assert.equal(results.filter((result) => result.status === 'fulfilled').length, 1);
});

void test('helper routes accept only explicit credentials/DTOs and sanitize responses and logs', async (t) => {
    const lines: string[] = [];
    const app = Fastify({ loggerInstance: pino({}, { write: (line: string) => { lines.push(line); } }) });
    const f = fixture();
    t.mock.method(helperService, 'claim', (ticket: string, ua: string) => f.service.claim(ticket, ua));
    t.mock.method(helperService, 'current', (cap: string, ua: string) => f.service.current(cap, ua));
    t.mock.method(helperService, 'password', (cap: string, ua: string, log: Parameters<typeof f.service.password>[2]) => f.service.password(cap, ua, log));
    await app.register(helperRoutes, { prefix: '/api/reauthorization-helper' });
    const { ticket } = await f.service.issue(ID, 1, UA, f.log);
    const headers = { host: 'outlook.wujiaqiao.dpdns.org', 'user-agent': UA };
    const claim = await app.inject({ method: 'POST', url: '/api/reauthorization-helper/claim', headers, payload: { ticket } });
    assert.equal(claim.statusCode, 200);
    const capability = (claim.json() as { data: { capability: string } }).data.capability;
    const capHeaders = { ...headers, 'x-reauthorization-capability': capability };
    for (const path of ['current', 'password']) {
        for (const extra of [{ authorization: 'Bearer management-jwt' }, { cookie: 'token=management-jwt' }, { host: 'evil.example' }]) {
            const response = await app.inject({ method: 'POST', url: `/api/reauthorization-helper/${path}`, headers: { ...capHeaders, ...extra }, payload: {} });
            assert.equal(response.statusCode, 401);
            assert.equal(response.headers['cache-control'], 'no-store');
        }
        assert.equal((await app.inject({ method: 'POST', url: `/api/reauthorization-helper/${path}`, headers, payload: {} })).statusCode, 400);
        assert.equal((await app.inject({ method: 'POST', url: `/api/reauthorization-helper/${path}`, headers: capHeaders, payload: { emailId: 8 } })).statusCode, 400);
    }
    const response = await app.inject({ method: 'POST', url: '/api/reauthorization-helper/password', headers: capHeaders, payload: {} });
    assert.equal(response.statusCode, 200);
    assert.equal(response.headers['cache-control'], 'no-store');
    assert.deepEqual((response.json() as { data: unknown }).data, { password: CLEAR });
    assert.ok(lines.some((line) => line.includes('helper_password_delivered')));
    t.mock.method(helperService, 'current', async () => { throw new Error(`raw-secret-${CLEAR}`); });
    const failed = await app.inject({ method: 'POST', url: '/api/reauthorization-helper/current', headers: capHeaders, payload: {} });
    assert.equal(failed.statusCode, 503);
    for (const secret of [CLEAR, ticket, capability, 'management-jwt', 'raw-secret']) {
        assert.equal(lines.some((line) => line.includes(secret)), false);
        assert.equal(failed.body.includes(secret), false);
    }
    assert.equal((await app.inject({ method: 'POST', url: '/api/reauthorization-helper/next', headers: capHeaders, payload: {} })).statusCode, 404);
    await app.close();
});

void test('ticket route requires JWT and SUPER_ADMIN; capability cannot authenticate management APIs', async (t) => {
    const app = Fastify();
    app.decorate('authenticateJwt', async (request) => {
        if (request.headers.authorization !== 'Bearer valid-admin-jwt') throw new AppError('UNAUTHORIZED', 'Authentication required', 401);
        request.user = { id: 1, username: 'test', role: request.headers['x-test-role'] === 'super' ? 'SUPER_ADMIN' : 'ADMIN' };
    });
    app.decorate('requireSuperAdmin', async (request) => {
        if (request.user?.role !== 'SUPER_ADMIN') throw new AppError('FORBIDDEN', 'Super admin required', 403);
    });
    t.mock.method(helperService, 'issue', async () => ({ ticket: 't'.repeat(43), expiresAt: new Date().toISOString() }));
    await app.register(adminRoutes, { prefix: '/admin/email-reauthorizations' });
    const url = `/admin/email-reauthorizations/${ID}/helper-ticket`;
    assert.equal((await app.inject({ method: 'POST', url, payload: {} })).statusCode, 401);
    assert.equal((await app.inject({ method: 'POST', url, headers: { authorization: 'Bearer valid-admin-jwt' }, payload: {} })).statusCode, 403);
    assert.equal((await app.inject({ method: 'POST', url, headers: { 'x-reauthorization-capability': 'c'.repeat(43) }, payload: {} })).statusCode, 401);
    const response = await app.inject({ method: 'POST', url, headers: { authorization: 'Bearer valid-admin-jwt', 'x-test-role': 'super' }, payload: {} });
    assert.equal(response.statusCode, 200);
    assert.equal(response.headers['cache-control'], 'no-store');
    await app.close();
});

void test('an unrelated remoteconnect tab cannot claim a globally visible ticket', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const result = await runMicrosoftScenario({
        url: 'https://login.live.com/oauth20_remoteconnect.srf',
        view: 'empty',
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            'gongxi-helper-ticket-v1': { ticket: 't'.repeat(43), runId, sessionId: ID, expiresAt: now + 60000 },
        },
    });
    assert.deepEqual(result.requests, []);
    assert.equal(result.clicks, 0);
    assert.equal('gongxiHelper' in result.tabState, false);
});

void test('an existing bound task cannot claim its ticket after unrelated navigation', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const boundAt = now - 1000;
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
        view: 'password',
        tabTask: { runId, sessionId: ID, bindingId, boundAt, expiresAt: now + 60000, actions: {} },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            [`gongxi-helper-ticket-v2:${bindingId}`]: { ticket: 't'.repeat(43), runId, sessionId: ID, bindingId,
                boundAt, publishedAt: now, expiresAt: now + 60000 },
        },
    });
    assert.deepEqual(result.requests, []);
    assert.equal(result.clicks, 0);
    assert.equal(result.values.has(`gongxi-helper-ticket-v2:${bindingId}`), false);
});

void test('a bound device tab claims once, saves observed submission proof, then fills a matching password page', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const control = { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false };
    const deviceResult = await runMicrosoftScenario({
        url: `https://microsoft.com/devicelogin#gongxi-helper=${bindingId}.${runId}.${ID}`,
        view: 'device',
        autoTicket: true,
        stopAfterClicks: 1,
        values: {
            'gongxi-helper-control-v1': control,
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(deviceResult.handoffEvents, ['binding-saved', 'ticket-published', 'ticket-claimed']);
    assert.deepEqual(deviceResult.requests, ['claim', 'current']);
    assert.equal(deviceResult.clicks, 1);
    assert.equal(deviceResult.submitEvents, 1);
    assert.equal(deviceResult.inputValues.device, 'ABCD-EFGH');
    const saved = (deviceResult.tabState as { gongxiHelper?: Record<string, unknown> }).gongxiHelper;
    assert.ok(saved);
    const proof = saved.deviceSubmission as { runId: string; sessionId: string; bindingId: string; clientId: string; submittedAt: number };
    assert.deepEqual({ ...proof, submittedAt: 0 }, { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: 0 });
    assert.equal(Number.isFinite(proof.submittedAt), true);
    assert.equal(saved.capability, 'c'.repeat(43));

    const passwordResult = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
        view: 'password',
        stopAfterClicks: 1,
        tabTask: saved,
        values: Object.fromEntries(deviceResult.values),
    });
    assert.deepEqual(passwordResult.requests, ['current', 'password']);
    assert.equal(passwordResult.clicks, 1);
    assert.equal(passwordResult.inputValues.password, CLEAR);
    const passwordTask = (passwordResult.tabState as { gongxiHelper: Record<string, unknown> }).gongxiHelper;
    assert.equal(passwordTask.passwordSent, true);
    assert.equal(passwordTask.identityVerified, true);
});

void test('a bound session still fills and submits a standard email page', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
        view: 'email',
        stopAfterClicks: 1,
        tabTask: {
            runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
            deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 1000 },
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests, ['current']);
    assert.equal(result.clicks, 1);
    assert.equal(result.inputValues.email, 'target@outlook.com');
});

void test('a password form whose submit button overrides POST with GET pauses before password claim', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
        view: 'password',
        formMethod: 'post',
        submitFormMethod: 'get',
        tabTask: {
            runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
            deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 1000 },
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests, ['current']);
    assert.equal(result.clicks, 0);
    assert.equal(result.inputValues.password, '');
});

for (const target of [
    { name: 'an external origin', url: 'https://evil.example/collect?client_id=client' },
    { name: 'an unrelated OAuth client', url: 'https://login.microsoftonline.com/common/login?client_id=unrelated-client' },
] as const) {
    void test(`a password submit button with formaction for ${target.name} pauses before password claim`, async () => {
        const now = Date.now();
        const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
        const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
        const result = await runMicrosoftScenario({
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
            view: 'password',
            formAction: 'https://login.microsoftonline.com/common/login?client_id=client',
            submitFormAction: target.url,
            tabTask: {
                runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
                deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 1000 },
            },
            values: {
                'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
                'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            },
        });
        assert.deepEqual(result.requests, ['current']);
        assert.equal(result.clicks, 0);
        assert.equal(result.inputValues.password, '');
    });
}

for (const mutation of [
    { name: 'page URL', value: { url: 'https://login.microsoftonline.com/common/changed?client_id=client' } },
    { name: 'form action', value: { formAction: 'https://login.microsoftonline.com/common/changed?client_id=client' } },
] as const) {
    void test(`a ${mutation.name} change during the first action save pauses before password claim`, async () => {
        const now = Date.now();
        const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
        const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
        const result = await runMicrosoftScenario({
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
            view: 'password',
            saveMutations: { 1: mutation.value },
            tabTask: {
                runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
                deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 1000 },
            },
            values: {
                'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
                'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            },
        });
        assert.deepEqual(result.requests, ['current']);
        assert.equal(result.clicks, 0);
        assert.equal(result.inputValues.password, '');
    });
}

void test('a target change while saving a filled password clears it and never clicks', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
        view: 'password',
        saveMutations: { 2: { submitFormAction: 'https://evil.example/collect?client_id=client' } },
        tabTask: {
            runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
            deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 1000 },
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests, ['current', 'password']);
    assert.equal(result.clicks, 0);
    assert.equal(result.inputValues.password, '');
});

for (const view of ['password-consent', 'password-continue'] as const) {
    void test(`a mixed ${view} page gives the manual-stop affordance priority over password`, async () => {
        const now = Date.now();
        const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
        const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
        const result = await runMicrosoftScenario({
            url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
            view,
            tabTask: {
                runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
                deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 1000 },
            },
            values: {
                'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
                'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            },
        });
        assert.deepEqual(result.requests, ['current']);
        assert.equal(result.clicks, 0);
        assert.equal(result.inputValues.password, '');
    });
}

for (const scenario of [
    { name: 'an unrelated OAuth client', url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=unrelated-client' },
    { name: 'a page without an explicit OAuth client', url: 'https://login.microsoftonline.com/common/login' },
] as const) {
    void test(`a matching password page for ${scenario.name} pauses before requesting the password`, async () => {
        const now = Date.now();
        const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
        const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
        const result = await runMicrosoftScenario({
            url: scenario.url,
            view: 'password',
            tabTask: {
                runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
                deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 1000 },
            },
            values: {
                'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
                'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            },
        });
        assert.deepEqual(result.requests, ['current']);
        assert.equal(result.clicks, 0);
        assert.equal(result.inputValues.password, '');
    });
}

void test('a device-submission proof with a tampered client id cannot release a password', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=client',
        view: 'password',
        tabTask: {
            runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
            deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'unrelated-client', submittedAt: now - 1000 },
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests, ['current']);
    assert.equal(result.clicks, 0);
    assert.equal(result.inputValues.password, '');
});

for (const view of ['password', 'consent'] as const) {
    void test(`a bound tab that has not submitted this session's device code cannot act on ${view}`, async () => {
        const now = Date.now();
        const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
        const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
        const result = await runMicrosoftScenario({
            url: 'https://login.live.com/oauth20_remoteconnect.srf',
            view,
            tabTask: { runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {} },
            values: {
                'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
                'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            },
        });
        assert.deepEqual(result.requests, ['current']);
        assert.equal(result.clicks, 0);
    });
}

void test('a submitted device-code proof cannot authorize consent for an unrelated OAuth client', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=unrelated-client',
        view: 'consent',
        stopAfterClicks: 1,
        tabTask: {
            runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
            identityVerified: true,
            deviceSubmission: { runId, sessionId: ID, bindingId, submittedAt: now - 1000 },
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.equal(result.clicks, 0, 'consent cannot be tied to the expected device-code application');
});

void test('userscript policy pauses challenges, unknown identities and stale controls; grants stay narrow', async () => {
    const source = await readFile(new URL('../../../../web/public/gongxi-mail-reauthorization.user.js', import.meta.url), 'utf8');
    const sandbox = { module: { exports: {} as {
        decide: (view: Record<string, unknown>, expectedEmail: string, expectedClientId: string, progress: Record<string, unknown>) => string;
        mayAct: (control: Record<string, unknown>, tab: Record<string, unknown>, now: number) => boolean;
        devicePath: (host: string, path: string) => boolean;
        workerRequiresPause: (control: Record<string, unknown>, worker: Record<string, unknown>, now: number) => boolean;
    } } };
    runInNewContext(source, sandbox, { timeout: 1000 });
    const policy = sandbox.module.exports;
    const deviceSubmission = { runId: 'r', sessionId: ID, bindingId: 'b', clientId: 'client', submittedAt: 1500 };
    const submitted = { runId: 'r', sessionId: ID, bindingId: 'b', deviceSubmission };
    assert.equal(policy.decide({ identities: [], password: true, clientIds: ['client'] }, 'target@outlook.com', 'client', {}), 'pause');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['client'] }, 'TARGET@outlook.com', 'client', {}), 'pause');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: [] }, 'TARGET@outlook.com', 'client', submitted), 'pause');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['unrelated'] }, 'TARGET@outlook.com', 'client', submitted), 'pause');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['client', 'client'] }, 'TARGET@outlook.com', 'client', submitted), 'password');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['client'] }, 'TARGET@outlook.com', 'client',
        { ...submitted, deviceSubmission: { ...deviceSubmission, clientId: 'tampered' } }), 'pause');
    assert.equal(policy.decide({ identities: ['other@outlook.com'], password: true, clientIds: ['client'] }, 'target@outlook.com', 'client', {}), 'mismatch');
    for (const field of ['challenge', 'error']) assert.equal(policy.decide({ identities: [], device: true, devicePath: true, [field]: true }, 'a@b.com', 'client', {}), 'pause');
    assert.equal(policy.decide({ identities: [], email: true, clientIds: ['client'] }, 'a@b.com', 'client', { ...submitted, identityVerified: true }), 'email');
    assert.equal(policy.decide({ identities: [], email: true, clientIds: ['client'] }, 'a@b.com', 'client', { identityVerified: true }), 'pause');
    for (const field of ['continue', 'consent']) {
        assert.equal(policy.decide({ identities: [], [field]: true, clientIds: ['client'] }, 'a@b.com', 'client', { ...submitted, identityVerified: true }), 'pause');
        assert.equal(policy.decide({ identities: ['a@b.com'], [field]: true, clientIds: ['client'] }, 'a@b.com', 'client', submitted), 'pause');
    }
    const control = { state: 'running', runId: 'r', sessionId: ID, bindingId: 'b', heartbeat: 1000, manual: false };
    const tab = { runId: 'r', sessionId: ID, bindingId: 'b', expiresAt: 50000 };
    assert.equal(policy.mayAct(control, tab, 2000), true);
    assert.equal(policy.mayAct({ ...control, manual: true }, tab, 2000), false);
    assert.equal(policy.mayAct(control, tab, 22000), false);
    assert.equal(policy.mayAct(control, { ...tab, sessionId: 'other' }, 2000), false);
    assert.equal(policy.mayAct(control, { ...tab, bindingId: 'other' }, 2000), false);
    assert.equal(policy.devicePath('login.microsoftonline.com', '/common/oauth2/deviceauth'), true);
    assert.equal(policy.devicePath('evil.example', '/common/oauth2/deviceauth'), false);
    assert.equal(policy.workerRequiresPause({ ...control, handoffAt: 1000 }, {}, 22000), true);
    assert.equal(policy.workerRequiresPause({ ...control, handoffAt: 1000 }, { ...tab, heartbeat: 2000, paused: true }, 2000), true);
    assert.equal(policy.workerRequiresPause({ ...control, handoffAt: 1000, manual: true }, { ...tab, heartbeat: 2000, paused: true }, 2000), false);
    assert.doesNotMatch(source, /@grant\s+none|unsafeWindow|@require\s|postMessage\(/);
    assert.deepEqual([...source.matchAll(/@connect\s+([^\s]+)/g)].map((match) => match[1]), ['outlook.wujiaqiao.dpdns.org']);
    assert.match(source, /cooldownUntil = Date\.now\(\) \+ 10000/);
    assert.match(source, /@sandbox\s+DOM/);
});
