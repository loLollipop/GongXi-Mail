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
const SECOND_ID = '87654321-4321-4321-8321-210987654321';
const UA = 'test-browser-agent';
const CLEAR = 'test-only-mailbox-credential';
const submittedLogin = (runId: string, bindingId: string, now: number) => ({
    deviceSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', submittedAt: now - 2000 },
    emailSubmission: { runId, sessionId: ID, bindingId, clientId: 'client', email: 'target@outlook.com', submittedAt: now - 1000 },
});

async function runMicrosoftScenario(options: {
    url: string;
    tabTask?: Record<string, unknown>;
    values?: Record<string, unknown>;
    view?: 'empty' | 'device' | 'picker' | 'password-choice' | 'email' | 'password' | 'stay' | 'consent' | 'password-consent' | 'password-continue';
    autoTicket?: boolean;
    stopAfterClicks?: number;
    formAction?: string;
    formMethod?: string;
    submitFormAction?: string;
    submitFormMethod?: string;
    currentClientId?: string;
    interruptSaveAfterClick?: boolean;
    suppressSubmitEvent?: boolean;
    clickView?: 'empty' | 'device' | 'picker' | 'password-choice' | 'email' | 'password' | 'stay' | 'consent' | 'password-consent' | 'password-continue';
    entryHref?: string;
    currentStatus?: 'PENDING' | 'SUCCEEDED';
    saveMutations?: Record<number, { url?: string; formAction?: string; formMethod?: string;
        submitFormAction?: string; submitFormMethod?: string }>;
}) {
    const source = await readFile(new URL('../../../../web/public/gongxi-mail-reauthorization.user.js', import.meta.url), 'utf8');
    const executable = source.replace(
        /[ ]{4}if \(location\.origin === ORIGIN\) \{[\s\S]*?\r?\n[ ]{4}\}\r?\n[ ]{4}else if \(MS_HOSTS\.includes\(location\.hostname\)\) void microsoftMain\(\);/,
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
    let activeView = options.view;
    const identity = makeNode(activeView === 'picker' ? 'cached@outlook.com' : 'target@outlook.com');
    const email = new TestInput();
    const password = new TestInput();
    const device = new TestInput();
    const submitText = ['consent', 'password-consent'].includes(options.view ?? '') ? 'Accept' :
        options.view === 'password-continue' || options.view === 'device' ? 'Continue' :
            options.view === 'email' ? 'Next' : 'Sign in';
    const submit = makeNode(submitText);
    const picker = makeNode('Use another account');
    const passwordOption = makeNode('Use your password');
    const no = makeNode('No');
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
        if (!options.suppressSubmitEvent && ['device', 'email', 'password'].includes(activeView ?? '')) {
            submitEvents++;
            for (const registered of [...submitListeners]) registered.listener();
            for (let index = submitListeners.length - 1; index >= 0; index--) {
                if (submitListeners[index].once) submitListeners.splice(index, 1);
            }
        }
        if (options.clickView) activeView = options.clickView;
        if (clicks >= (options.stopAfterClicks ?? Number.POSITIVE_INFINITY)) clickReached?.();
        },
    });
    Object.assign(picker, {
        click: () => {
            clicks++;
            if (clicks >= (options.stopAfterClicks ?? Number.POSITIVE_INFINITY)) clickReached?.();
        },
    });
    Object.assign(passwordOption, {
        href: options.entryHref ?? '',
        getAttribute: (name: string) => name === 'href' ? options.entryHref ?? null : null,
        click: () => {
            clicks++;
            if (clicks >= (options.stopAfterClicks ?? Number.POSITIVE_INFINITY)) clickReached?.();
        },
    });
    Object.assign(no, {
        form,
        formAction: '',
        formMethod: '',
        getAttribute: () => null,
        closest: () => form,
        click: () => {
            clicks++;
            if (clicks >= (options.stopAfterClicks ?? Number.POSITIVE_INFINITY)) clickReached?.();
        },
    });
    const bodyText = () => ['consent', 'password-consent'].includes(activeView ?? '') ? 'Permissions requested by this app' :
        activeView === 'password-continue' ? 'Continue signing in to this app' : activeView === 'stay' ? 'Stay signed in?' : '';
    const querySelectorAll = (selector: string) => {
        if (selector.startsWith('#displayName,')) return ['picker', 'password', 'stay', 'consent', 'password-consent', 'password-continue'].includes(activeView ?? '') ? [identity] : [];
        if (selector === '#i0116[name="loginfmt"], #usernameEntry[type="email"][autocomplete~="username"]') return activeView === 'email' ? [email] : [];
        if (selector === 'input#i0118[name="passwd"][type="password"], #passwordEntry[type="password"], input[type="password"][autocomplete="current-password"]') return ['password', 'password-consent', 'password-continue'].includes(activeView ?? '') ? [password] : [];
        if (selector === '#otc, input[name="user_code"]') return activeView === 'device' ? [device] : [];
        if (selector === '#idSIButton9, #idSubmit_Consent, #idBtn_Accept, form button[type="submit"]') return ['empty', 'picker'].includes(activeView ?? '') ? [] : [submit];
        if (selector === '#otherTile, #idDiv_UseAnotherAccount') return activeView === 'picker' ? [picker] : [];
        if (selector === 'button, a, [role="button"]') return activeView === 'password-choice' ? [passwordOption] : [];
        if (selector === 'button, input[type="button"], input[type="submit"]') return activeView === 'stay' ? [no] : [];
        if (selector.startsWith('#idDiv_SAOTCS_Proofs,')) return activeView === 'password-choice' ? [makeNode('verification')] : [];
        return [];
    };
    const makeElement = () => ({
        id: '', style: { cssText: '' }, textContent: '',
        append: () => undefined, addEventListener: () => undefined,
        attachShadow: () => ({ append: () => undefined }),
    });
    const documentValue = {
        body: { get innerText() { return bodyText(); }, append: () => undefined },
        createElement: makeElement,
        querySelectorAll,
    };
    const windowValue: Record<string, unknown> = { addEventListener: () => undefined };
    windowValue.top = windowValue;
    windowValue.self = windowValue;
    const tabState = options.tabTask ? { gongxiHelper: { ...options.tabTask } } : {};
    let durableTabState = options.interruptSaveAfterClick ? structuredClone(tabState) : tabState;
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
            return (value: object, done?: () => void) => {
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
                if (options.interruptSaveAfterClick && clicks > 0) return;
                if (options.interruptSaveAfterClick) durableTabState = structuredClone(value);
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
                    sessionId: ID, email: 'target@outlook.com', userCode: 'ABCD-EFGH', status: options.currentStatus ?? 'PENDING',
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
        handoffEvents, tabState: durableTabState, values };
}

async function runAdminManualSuccessScenario(options: {
    outOfOrderSessionRead?: boolean;
    crossSessionDoubleContinue?: boolean;
    exerciseStaleResponseGuard?: boolean;
    pauseDuringContinue?: boolean;
} = {}) {
    const source = await readFile(new URL('../../../../web/public/gongxi-mail-reauthorization.user.js', import.meta.url), 'utf8');
    let executable = source.replace(
        /[ ]{4}if \(location\.origin === ORIGIN\) \{[\s\S]*?\r?\n[ ]{4}\}\r?\n[ ]{4}else if \(MS_HOSTS\.includes\(location\.hostname\)\) void microsoftMain\(\);/,
        '    module.exports.adminMain = adminMain;',
    );
    assert.notEqual(executable, source, 'admin userscript test hook must replace the runtime dispatch');
    if (options.exerciseStaleResponseGuard) {
        const guarded = executable;
        executable = executable.replace('                if (continuePending) return;\n                continuePending = true;',
            '                continuePending = true;');
        assert.notEqual(executable, guarded, 'the test must bypass only the click lock to exercise stale-response isolation');
    }
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const secondBindingId = 'cccccccc-cccc-4ccc-8ccc-cccccccccccc';
    let scenarioNow = Date.now();
    class ScenarioDate extends Date {
        constructor(value?: string | number) {
            if (value === undefined) super(scenarioNow);
            else super(value);
        }
        static override now() { return scenarioNow; }
    }
    const buttons = new Map<string, () => void>();
    const makeElement = () => {
        const element = {
            id: '', style: { cssText: '' }, textContent: '',
            append: (..._children: unknown[]) => undefined,
            attachShadow: () => ({ append: (..._children: unknown[]) => undefined }),
            addEventListener: (type: string, listener: () => void) => {
                if (type === 'click') buttons.set(element.textContent, listener);
            },
        };
        return element;
    };
    const values = new Map<string, unknown>([['gongxi-helper-worker-v1', {
        runId, sessionId: ID, bindingId, heartbeat: scenarioNow, paused: true, freshLogin: false, reason: '人工验证',
    }]]);
    const controlHistory: Record<string, unknown>[] = [];
    let candidateReads = 0;
    let sessionReads = 0;
    let secondSessionReads = 0;
    let ticketReads = 0;
    let releaseStaleSessionRead: (() => void) | undefined;
    let releaseLateSucceededRead: (() => void) | undefined;
    let lateSucceededReadResolved = false;
    let openedTab: { closed: boolean; close: () => void } | null = null;
    const session = (status: 'PENDING' | 'SUCCEEDED', sessionId = ID) => ({
        sessionId, email: sessionId === ID ? 'target@outlook.com' : 'second@outlook.com', status, clientId: 'client',
        verificationUri: 'https://www.microsoft.com/link',
        nextPollAt: new ScenarioDate(scenarioNow + 60000).toISOString(), serverTime: new ScenarioDate().toISOString(),
    });
    const moduleValue: { exports: { adminMain?: () => Promise<void> } } = { exports: {} };
    const sandbox = {
        module: moduleValue,
        document: {
            body: { append: (..._children: unknown[]) => undefined },
            createElement: makeElement,
            documentElement: {
                setAttribute: (_name: string, _value: string) => undefined,
                removeAttribute: (_name: string) => undefined,
            },
        },
        window: { top: null, self: null, addEventListener: (_type: string, _listener: () => void) => undefined },
        location: { origin: 'https://outlook.wujiaqiao.dpdns.org', protocol: 'https:',
            hostname: 'outlook.wujiaqiao.dpdns.org', pathname: '/reauthorizations', hash: '',
            href: 'https://outlook.wujiaqiao.dpdns.org/reauthorizations' },
        history: { replaceState: () => undefined },
        navigator: { locks: { request: async (_name: string, _options: object,
            action: (lock: object) => Promise<void>) => action({}) } },
        crypto: { randomUUID: (() => {
            const ids = [runId, bindingId, secondBindingId];
            return () => ids.shift() ?? 'dddddddd-dddd-4ddd-8ddd-dddddddddddd';
        })() },
        localStorage: { getItem: (key: string) => key === 'token' ? 'admin-jwt' : null },
        URL, URLSearchParams, btoa, atob, JSON, Date: ScenarioDate,
        setTimeout: (callback: (...args: unknown[]) => void, delay = 0, ...args: unknown[]) => {
            scenarioNow += Math.max(0, delay);
            return setTimeout(callback, 0, ...args);
        },
        clearTimeout,
        setInterval: () => 1,
        clearInterval: (_handle: number) => undefined,
        fetch: async (raw: string, init?: { method?: string }) => {
            const url = new URL(raw);
            let data: unknown;
            if (url.pathname.endsWith('/candidates')) {
                candidateReads++;
                data = candidateReads === 1 ? [{ supported: true, emailId: 7, activeSession: session('PENDING') }] :
                    (options.crossSessionDoubleContinue || options.exerciseStaleResponseGuard) && candidateReads === 2
                        ? [{ supported: true, emailId: 8, activeSession: session('PENDING', SECOND_ID) }]
                        : [];
            } else if (url.pathname.endsWith('/helper-ticket') && init?.method === 'POST') {
                ticketReads++;
                data = { ticket: 't'.repeat(43), expiresAt: new ScenarioDate(scenarioNow + 60000).toISOString() };
            } else if (url.pathname.endsWith(`/${ID}`)) {
                sessionReads++;
                if ((options.outOfOrderSessionRead || options.crossSessionDoubleContinue || options.pauseDuringContinue ||
                    options.exerciseStaleResponseGuard) && sessionReads === 2) {
                    data = await new Promise<ReturnType<typeof session>>((resolve) => {
                        releaseStaleSessionRead = () => resolve(session('PENDING'));
                    });
                } else if ((options.pauseDuringContinue && sessionReads === 3) ||
                    (options.exerciseStaleResponseGuard && sessionReads === 4)) {
                    data = await new Promise<ReturnType<typeof session>>((resolve) => {
                        releaseLateSucceededRead = () => {
                            lateSucceededReadResolved = true;
                            resolve(session('SUCCEEDED'));
                        };
                    });
                } else {
                    data = session(sessionReads === 1 ? 'PENDING' : 'SUCCEEDED');
                }
            } else if (url.pathname.endsWith(`/${SECOND_ID}`)) {
                secondSessionReads++;
                const secondWorker = values.get('gongxi-helper-worker-v1') as Record<string, unknown> | undefined;
                if (secondWorker?.sessionId === SECOND_ID) {
                    values.set('gongxi-helper-worker-v1', { ...secondWorker, heartbeat: scenarioNow });
                }
                data = session('PENDING', SECOND_ID);
                if (options.exerciseStaleResponseGuard && secondSessionReads === 1) {
                    setTimeout(() => releaseLateSucceededRead?.(), 0);
                }
            } else {
                throw new Error(`unexpected admin request ${url.pathname}`);
            }
            return { ok: true, json: async () => ({ success: true, data }) };
        },
        GM_getValue: (key: string, fallback: unknown) => {
            if (values.has(key)) return values.get(key);
            if (key.startsWith('gongxi-helper-binding-v2:')) {
                const control = values.get('gongxi-helper-control-v1') as Record<string, unknown> | undefined;
                const requestedBindingId = key.slice('gongxi-helper-binding-v2:'.length);
                if (control?.runId === runId && control.bindingId === requestedBindingId) {
                    return { runId, sessionId: control.sessionId, bindingId: requestedBindingId, boundAt: control.handoffAt };
                }
            }
            return fallback;
        },
        GM_setValue: (key: string, value: unknown) => {
            values.set(key, value);
            if (key === 'gongxi-helper-control-v1') {
                const control = value as Record<string, unknown>;
                controlHistory.push(structuredClone(control));
                if ((options.crossSessionDoubleContinue || options.exerciseStaleResponseGuard) &&
                    control.sessionId === SECOND_ID && control.bindingId === secondBindingId) {
                    values.set('gongxi-helper-worker-v1', {
                        runId, sessionId: SECOND_ID, bindingId: secondBindingId,
                        heartbeat: Math.max(scenarioNow, control.handoffAt as number),
                        paused: false, freshLogin: true, freshLoginAt: control.handoffAt,
                    });
                }
            }
        },
        GM_deleteValue: (key: string) => { values.delete(key); },
        GM_openInTab: () => {
            openedTab = { closed: false, close() { this.closed = true; } };
            return openedTab;
        },
    };
    Object.assign(sandbox.window, { top: sandbox.window, self: sandbox.window });
    runInNewContext(executable, sandbox, { timeout: 1000 });
    assert.ok(moduleValue.exports.adminMain);
    await moduleValue.exports.adminMain();
    const start = buttons.get('启动 / 继续队列');
    const pause = buttons.get('暂停（当前项转人工）');
    const finish = buttons.get('结束助手 / 切换人工');
    assert.ok(start && pause && finish);
    start();
    const waitFor = async (predicate: () => boolean) => {
        for (let attempts = 0; attempts < 200; attempts++) {
            if (predicate()) return;
            await new Promise((resolve) => setTimeout(resolve, 5));
        }
        assert.fail('admin userscript scenario timed out');
    };
    await waitFor(() => {
        const control = values.get('gongxi-helper-control-v1') as Record<string, unknown> | undefined;
        return control?.state === 'paused' && control.manual === true;
    });
    if (options.outOfOrderSessionRead || options.crossSessionDoubleContinue || options.exerciseStaleResponseGuard ||
        options.pauseDuringContinue) {
        await waitFor(() => sessionReads === 2);
    }
    const generationBeforeContinue = (values.get('gongxi-helper-control-v1') as { generation?: number }).generation;
    start();
    if (options.pauseDuringContinue) {
        await waitFor(() => !!releaseLateSucceededRead);
        pause();
        releaseLateSucceededRead?.();
        releaseStaleSessionRead?.();
        await new Promise((resolve) => setTimeout(resolve, 50));
        const controlBeforeFinish = structuredClone(values.get('gongxi-helper-control-v1') as Record<string, unknown>);
        finish();
        return { bindingId, secondBindingId, candidateReads, sessionReads, secondSessionReads, ticketReads,
            controlHistory, openedTab, controlBeforeFinish, generationBeforeContinue,
            lateSucceededReadWasIssued: true, lateSucceededReadResolved };
    }
    if (options.crossSessionDoubleContinue || options.exerciseStaleResponseGuard) {
        start();
        await new Promise((resolve) => setTimeout(resolve, 25));
    }
    await waitFor(() => controlHistory.some((control) => {
        const proof = control.manualSuccess as Record<string, unknown> | undefined;
        return control.state === 'running' && control.manual === false && proof?.sessionId === ID;
    }));
    const proofIndex = controlHistory.findIndex((control) => {
        const proof = control.manualSuccess as Record<string, unknown> | undefined;
        return control.state === 'running' && control.manual === false && proof?.sessionId === ID;
    });
    if (options.outOfOrderSessionRead || options.crossSessionDoubleContinue || options.exerciseStaleResponseGuard) {
        const release = releaseStaleSessionRead;
        assert.ok(release);
        release();
    }
    await waitFor(() => controlHistory.slice(proofIndex + 1).some((control) =>
        control.sessionId === null && control.bindingId === null && control.manualSuccess === null));
    const cleanup = controlHistory.slice(proofIndex + 1).find((control) =>
        control.sessionId === null && control.bindingId === null && control.manualSuccess === null);
    if (cleanup?.state === 'running') await waitFor(() => candidateReads === 2);
    let controlBeforeFinish: Record<string, unknown> | undefined;
    if (options.crossSessionDoubleContinue || options.exerciseStaleResponseGuard) {
        await waitFor(() => ticketReads === 2 && secondSessionReads > 0);
        await new Promise((resolve) => setTimeout(resolve, 50));
        controlBeforeFinish = structuredClone(values.get('gongxi-helper-control-v1') as Record<string, unknown>);
    }
    finish();
    return { bindingId, secondBindingId, candidateReads, sessionReads, secondSessionReads, ticketReads,
        controlHistory, openedTab, controlBeforeFinish, generationBeforeContinue,
        lateSucceededReadWasIssued: !!releaseLateSucceededRead,
        lateSucceededReadResolved };
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

void test('a bound device tab claims once, submits a fresh account, then fills a matching password page without client_id URLs', async () => {
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

    const emailResult = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/login',
        view: 'email',
        stopAfterClicks: 1,
        tabTask: saved,
        values: Object.fromEntries(deviceResult.values),
    });
    assert.deepEqual(emailResult.requests, ['current']);
    assert.equal(emailResult.clicks, 1);
    assert.equal(emailResult.inputValues.email, 'target@outlook.com');
    const emailTask = (emailResult.tabState as { gongxiHelper: Record<string, unknown> }).gongxiHelper;
    assert.deepEqual({ ...(emailTask.emailSubmission as object), submittedAt: 0 }, {
        runId, sessionId: ID, bindingId, clientId: 'client', email: 'target@outlook.com', submittedAt: 0,
    });

    const passwordResult = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/login',
        view: 'password',
        stopAfterClicks: 1,
        tabTask: emailTask,
        values: Object.fromEntries(emailResult.values),
    });
    assert.deepEqual(passwordResult.requests, ['current', 'password']);
    assert.equal(passwordResult.clicks, 1);
    assert.equal(passwordResult.inputValues.password, CLEAR);
    const passwordTask = (passwordResult.tabState as { gongxiHelper: Record<string, unknown> }).gongxiHelper;
    assert.equal(passwordTask.passwordSent, true);
    assert.equal(passwordTask.identityVerified, true);
});

void test('device and email submit proofs survive navigation interrupting GM_saveTab callbacks', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const boundAt = now - 1000;
    const control = { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false, handoffAt: boundAt };
    const values = {
        'gongxi-helper-control-v1': control,
        'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
    };
    const deviceResult = await runMicrosoftScenario({
        url: 'https://www.microsoft.com/link',
        view: 'device',
        stopAfterClicks: 1,
        interruptSaveAfterClick: true,
        suppressSubmitEvent: true,
        tabTask: { runId, sessionId: ID, bindingId, boundAt, expiresAt: now + 60000,
            capability: 'c'.repeat(43), actions: {} },
        values,
    });
    const deviceTask = (deviceResult.tabState as { gongxiHelper: Record<string, unknown> }).gongxiHelper;
    assert.equal(deviceTask.deviceSubmission, undefined, 'the simulated navigation must discard the tab save');
    const deviceReceipt = deviceResult.values.get('gongxi-helper-submission-v1') as Record<string, unknown>;
    assert.deepEqual({ runId: deviceReceipt.runId, sessionId: deviceReceipt.sessionId,
        bindingId: deviceReceipt.bindingId, clientId: deviceReceipt.clientId },
        { runId, sessionId: ID, bindingId, clientId: 'client' });
    assert.equal(JSON.stringify(deviceReceipt).includes('c'.repeat(43)), false, 'the capability must remain tab-private');
    assert.equal(JSON.stringify(deviceReceipt).includes(CLEAR), false, 'the password must never enter shared proof storage');

    const emailResult = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/login',
        view: 'email',
        stopAfterClicks: 1,
        interruptSaveAfterClick: true,
        tabTask: deviceTask,
        values: Object.fromEntries(deviceResult.values),
    });
    assert.equal(emailResult.clicks, 1, 'the next document must recover the device submit proof');
    const emailTask = (emailResult.tabState as { gongxiHelper: Record<string, unknown> }).gongxiHelper;
    assert.equal(emailTask.emailSubmission, undefined, 'the simulated navigation must discard the email tab save');
    const emailReceipt = emailResult.values.get('gongxi-helper-submission-v1') as { emailSubmission?: unknown };
    assert.ok(emailReceipt.emailSubmission, 'the verified email-button invocation must synchronously replace the durable receipt');

    const passwordResult = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/login',
        view: 'password',
        stopAfterClicks: 1,
        interruptSaveAfterClick: true,
        tabTask: emailTask,
        values: Object.fromEntries(emailResult.values),
    });
    assert.deepEqual(passwordResult.requests, ['current', 'password']);
    assert.equal(passwordResult.clicks, 1, 'the password boundary must require the recovered email submit proof');
    assert.equal(passwordResult.inputValues.password, CLEAR);
    const passwordTask = (passwordResult.tabState as { gongxiHelper: Record<string, unknown> }).gongxiHelper;
    assert.equal(passwordTask.passwordSubmission, undefined, 'navigation must be able to interrupt the tab save after password click');
    const passwordWorker = passwordResult.values.get('gongxi-helper-worker-v1') as { freshLogin?: boolean; freshLoginAt?: number };
    assert.equal(passwordWorker.freshLogin, true);
    assert.equal(Number.isFinite(passwordWorker.freshLoginAt), true);

    const stayResult = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/login',
        view: 'stay',
        stopAfterClicks: 1,
        tabTask: passwordTask,
        values: Object.fromEntries(passwordResult.values),
    });
    assert.equal(stayResult.clicks, 1);
    const inheritedWorker = stayResult.values.get('gongxi-helper-worker-v1') as { freshLogin?: boolean; freshLoginAt?: number };
    assert.equal(inheritedWorker.freshLogin, true);
    assert.equal(inheritedWorker.freshLoginAt, passwordWorker.freshLoginAt,
        'the next document must retain the validated password-click timestamp');
});

void test('a local password submission restores fresh-login proof only for the bound, post-handoff, non-future click', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const handoffAt = now - 2000;
    const run = async (passwordSubmission: Record<string, unknown>) => {
        const result = await runMicrosoftScenario({
            url: 'https://login.microsoftonline.com/common/login',
            view: 'empty',
            currentStatus: 'SUCCEEDED',
            tabTask: {
                runId, sessionId: ID, bindingId, boundAt: handoffAt, expiresAt: now + 60000,
                capability: 'c'.repeat(43), actions: {}, passwordSubmission,
            },
            values: {
                'gongxi-helper-control-v1': {
                    state: 'running', runId, sessionId: ID, bindingId, heartbeat: now,
                    manual: false, handoffAt,
                },
                'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
            },
        });
        return result.values.get('gongxi-helper-worker-v1') as { freshLogin?: boolean; freshLoginAt?: number };
    };
    const submission = {
        runId, sessionId: ID, bindingId, clientId: 'client', email: 'target@outlook.com', submittedAt: now - 1000,
    };
    const restored = await run(submission);
    assert.equal(restored.freshLogin, true, 'the tab-local click proof must survive a missing shared worker record');
    assert.equal(restored.freshLoginAt, submission.submittedAt);
    assert.equal((await run({ ...submission, bindingId: 'cccccccc-cccc-4ccc-8ccc-cccccccccccc' })).freshLogin, false,
        'a local proof from another binding must be rejected');
    assert.equal((await run({ ...submission, submittedAt: handoffAt - 1 })).freshLogin, false,
        'a local proof from before this handoff must be rejected');
    assert.equal((await run({ ...submission, submittedAt: Date.now() + 60000 })).freshLogin, false,
        'a future local proof must be rejected');
});

void test('admin manual success advances once without a stale worker restoring paused state', async () => {
    const result = await runAdminManualSuccessScenario();
    const proofIndex = result.controlHistory.findIndex((control) => {
        const proof = control.manualSuccess as Record<string, unknown> | undefined;
        return control.state === 'running' && control.manual === false && proof?.sessionId === ID;
    });
    assert.notEqual(proofIndex, -1, 'the continue click must record the backend-confirmed manual success');
    const proofControl = result.controlHistory[proofIndex];
    assert.deepEqual(proofControl.manualSuccess, {
        runId: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa', sessionId: ID,
        bindingId: result.bindingId, confirmedAt: (proofControl.manualSuccess as { confirmedAt: number }).confirmedAt,
    });
    assert.equal(Number.isFinite((proofControl.manualSuccess as { confirmedAt: number }).confirmedAt), true);
    const cleanupIndex = result.controlHistory.findIndex((control, index) => index > proofIndex &&
        control.state === 'running' && control.sessionId === null && control.bindingId === null && control.manualSuccess === null);
    assert.notEqual(cleanupIndex, -1, 'the verified item must be cleared so the queue can select the next item');
    assert.equal(result.controlHistory.slice(proofIndex + 1, cleanupIndex).some((control) => control.state === 'paused'), false,
        'the stale paused worker must not overwrite the verified manual-success transition');
    assert.equal(result.ticketReads, 1, 'manual recovery must not issue another helper ticket');
    assert.ok(result.sessionReads >= 2, 'the backend session must be observed as SUCCEEDED before cleanup');
    assert.equal(result.candidateReads, 2, 'the queue must automatically ask for the next candidate after cooldown');
});

void test('pausing while manual-success verification is in flight invalidates the late success', async () => {
    const result = await runAdminManualSuccessScenario({ pauseDuringContinue: true });
    assert.equal(result.lateSucceededReadResolved, true, 'the delayed success must actually arrive after the pause');
    assert.equal(result.controlBeforeFinish?.state, 'paused');
    assert.equal(result.controlBeforeFinish?.manual, true);
    assert.equal(result.controlBeforeFinish?.sessionId, ID);
    assert.equal(result.controlBeforeFinish?.bindingId, result.bindingId);
    assert.equal(result.controlBeforeFinish?.manualSuccess, null);
    assert.ok((result.controlBeforeFinish?.generation as number) > (result.generationBeforeContinue ?? -1),
        'pause must advance and publish the generation that guards in-flight requests');
    assert.equal(result.candidateReads, 1, 'the late success must not advance from A to candidate B');
    assert.equal(result.secondSessionReads, 0, 'candidate B must never start');
    assert.equal(result.ticketReads, 1, 'the late success must not issue a helper ticket for B');
});

void test('admin manual success ignores an older pending response that arrives out of order', async () => {
    const result = await runAdminManualSuccessScenario({ outOfOrderSessionRead: true });
    const proofIndex = result.controlHistory.findIndex((control) => {
        const proof = control.manualSuccess as Record<string, unknown> | undefined;
        return control.state === 'running' && control.manual === false && proof?.sessionId === ID;
    });
    assert.notEqual(proofIndex, -1);
    const cleanupIndex = result.controlHistory.findIndex((control, index) => index > proofIndex &&
        control.state === 'running' && control.sessionId === null && control.bindingId === null && control.manualSuccess === null);
    assert.notEqual(cleanupIndex, -1, 'the late PENDING response must not leave the cleaned queue paused');
    assert.equal(result.controlHistory.slice(proofIndex + 1, cleanupIndex).some((control) => control.state === 'paused'), false);
    assert.equal(result.candidateReads, 2, 'one continue click must advance to the next candidate');
});

void test('double continue cannot let a late succeeded response for the prior session clear the active next session', async () => {
    const result = await runAdminManualSuccessScenario({ crossSessionDoubleContinue: true });
    assert.equal(result.lateSucceededReadWasIssued, false,
        'the second continue click must not start another verification request while the first is pending');
    assert.deepEqual(result.controlBeforeFinish, {
        runId: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa', state: 'running', sessionId: SECOND_ID,
        bindingId: result.secondBindingId, manual: false,
        heartbeat: result.controlBeforeFinish?.heartbeat,
        generation: result.controlBeforeFinish?.generation,
        handoffAt: result.controlBeforeFinish?.handoffAt,
        manualSuccess: null,
        reason: '',
    });
    assert.equal(result.candidateReads, 2, 'the active second session must not be cleaned or advance the queue');
    assert.ok(result.secondSessionReads > 0, 'the second session must have actually entered backend polling');
});

void test('a delayed succeeded response for the prior session is discarded after the next session becomes current', async () => {
    const result = await runAdminManualSuccessScenario({ exerciseStaleResponseGuard: true });
    assert.equal(result.lateSucceededReadWasIssued, true);
    assert.equal(result.lateSucceededReadResolved, true);
    assert.equal(result.controlBeforeFinish?.sessionId, SECOND_ID);
    assert.equal(result.controlBeforeFinish?.bindingId, result.secondBindingId);
    assert.equal(result.controlBeforeFinish?.state, 'running');
    assert.equal(result.controlBeforeFinish?.manual, false);
    assert.equal(result.candidateReads, 2, 'the stale success must not clear B or advance to another candidate');
});

void test('a bound session still fills and submits a standard email page', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/login',
        view: 'email',
        stopAfterClicks: 1,
        tabTask: {
            runId, sessionId: ID, bindingId, boundAt: now - 2000, expiresAt: now + 60000,
            capability: 'c'.repeat(43), actions: { device: true },
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

void test('an account picker with cached identities chooses use another account without requiring a form', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/oauth2/authorize',
        view: 'picker',
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
});

void test('a passwordless-first verification page chooses the explicit password alternative', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.live.com/oauth20_remoteconnect.srf',
        view: 'password-choice',
        stopAfterClicks: 1,
        tabTask: {
            runId, sessionId: ID, bindingId, boundAt: now - 3000, expiresAt: now + 60000,
            capability: 'c'.repeat(43), actions: { device: true, email: true },
            ...submittedLogin(runId, bindingId, now),
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests, ['current']);
    assert.equal(result.clicks, 1);
});

void test('a password alternative with an untrusted entry target pauses', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.live.com/oauth20_remoteconnect.srf',
        view: 'password-choice',
        entryHref: 'https://evil.example/collect',
        tabTask: {
            runId, sessionId: ID, bindingId, boundAt: now - 3000, expiresAt: now + 60000,
            capability: 'c'.repeat(43), actions: { device: true, email: true },
            ...submittedLogin(runId, bindingId, now),
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests, ['current']);
    assert.equal(result.clicks, 0);
});

void test('the current stay-signed-in layout chooses the explicit negative option', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.live.com/ppsecure/post.srf',
        view: 'stay',
        stopAfterClicks: 1,
        tabTask: {
            runId, sessionId: ID, bindingId, boundAt: now - 4000, expiresAt: now + 60000,
            capability: 'c'.repeat(43), passwordSent: true,
            actions: { device: true, email: true, password: true },
            ...submittedLogin(runId, bindingId, now),
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests, ['current']);
    assert.equal(result.clicks, 1);
});

void test('a cached matching password page pauses until this bound flow submitted the queued email', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.live.com/oauth20_remoteconnect.srf',
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
            ...submittedLogin(runId, bindingId, now),
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
                ...submittedLogin(runId, bindingId, now),
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
                ...submittedLogin(runId, bindingId, now),
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
            ...submittedLogin(runId, bindingId, now),
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

void test('a password click that does not navigate is cleared when the resulting page pauses', async () => {
    const now = Date.now();
    const runId = 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa';
    const bindingId = 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb';
    const result = await runMicrosoftScenario({
        url: 'https://login.microsoftonline.com/common/login',
        view: 'password',
        clickView: 'consent',
        tabTask: {
            runId, sessionId: ID, bindingId, expiresAt: now + 60000, capability: 'c'.repeat(43), actions: {},
            ...submittedLogin(runId, bindingId, now),
        },
        values: {
            'gongxi-helper-control-v1': { state: 'running', runId, sessionId: ID, bindingId, heartbeat: now, manual: false },
            'gongxi-helper-heartbeat-v1': { runId, heartbeat: now },
        },
    });
    assert.deepEqual(result.requests.slice(0, 3), ['current', 'password', 'current']);
    assert.equal(result.clicks, 1);
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
                ...submittedLogin(runId, bindingId, now),
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
                ...submittedLogin(runId, bindingId, now),
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
    const sandbox = { URL, URLSearchParams, btoa, atob, JSON, module: { exports: {} as {
        decide: (view: Record<string, unknown>, expectedEmail: string, expectedClientId: string, progress: Record<string, unknown>) => string;
        mayAct: (control: Record<string, unknown>, tab: Record<string, unknown>, now: number) => boolean;
        devicePath: (host: string, path: string) => boolean;
        verificationTarget: (session: Record<string, unknown>, fragment: string) => string | null;
        bootstrapTarget: (session: Record<string, unknown>, bindingId: string, runId: string) => string | null;
        recoverSubmissionProofs: (stored: Record<string, unknown>, progress: Record<string, unknown>,
            expectedEmail: string, expectedClientId: string, now: number) => Record<string, unknown> | null;
        workerRequiresPause: (control: Record<string, unknown>, worker: Record<string, unknown>, now: number) => boolean;
        workerConfirmedFreshLogin: (control: Record<string, unknown>, worker: Record<string, unknown>, now: number) => boolean;
        manualSuccessConfirmed: (control: Record<string, unknown>, session: Record<string, unknown>, now: number) => boolean;
    } } };
    runInNewContext(source, sandbox, { timeout: 1000 });
    const policy = sandbox.module.exports;
    const deviceSubmission = { runId: 'r', sessionId: ID, bindingId: 'b', clientId: 'client', submittedAt: 1500 };
    const submitted = { runId: 'r', sessionId: ID, bindingId: 'b', deviceSubmission,
        emailSubmission: { runId: 'r', sessionId: ID, bindingId: 'b', clientId: 'client', email: 'target@outlook.com', submittedAt: 1600 } };
    assert.equal(policy.decide({ identities: [], password: true, clientIds: ['client'] }, 'target@outlook.com', 'client', {}), 'pause');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['client'] }, 'TARGET@outlook.com', 'client', {}), 'pause');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: [] }, 'TARGET@outlook.com', 'client', submitted), 'password');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['unrelated'] }, 'TARGET@outlook.com', 'client', submitted), 'pause');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['client', 'client'] }, 'TARGET@outlook.com', 'client', submitted), 'password');
    assert.equal(policy.decide({ identities: ['target@outlook.com'], password: true, clientIds: ['client'] }, 'TARGET@outlook.com', 'client',
        { ...submitted, deviceSubmission: { ...deviceSubmission, clientId: 'tampered' } }), 'pause');
    assert.equal(policy.decide({ identities: ['other@outlook.com'], password: true, clientIds: ['client'] }, 'target@outlook.com', 'client', {}), 'mismatch');
    assert.equal(policy.decide({ identities: ['cached@outlook.com'], picker: true, clientIds: [] }, 'target@outlook.com', 'client',
        { ...submitted, emailSubmission: undefined }), 'choose-other');
    for (const field of ['challenge', 'error']) assert.equal(policy.decide({ identities: [], device: true, devicePath: true, [field]: true }, 'a@b.com', 'client', {}), 'pause');
    assert.equal(policy.decide({ identities: [], email: true, clientIds: ['client'] }, 'a@b.com', 'client', { ...submitted, identityVerified: true }), 'email');
    assert.equal(policy.decide({ identities: [], email: true, clientIds: ['client'] }, 'a@b.com', 'client', { identityVerified: true }), 'pause');
    for (const field of ['continue', 'consent']) {
        assert.equal(policy.decide({ identities: [], [field]: true, clientIds: ['client'] }, 'a@b.com', 'client', { ...submitted, identityVerified: true }), 'pause');
        assert.equal(policy.decide({ identities: ['a@b.com'], [field]: true, clientIds: ['client'] }, 'a@b.com', 'client', submitted), 'pause');
    }
    const proofTask = { ...submitted, boundAt: 1000, expiresAt: 5000, actions: { device: true, email: true } };
    const storedProof = { version: 1, runId: 'r', sessionId: ID, bindingId: 'b', clientId: 'client',
        email: 'target@outlook.com', boundAt: 1000, expiresAt: 5000, deviceSubmission,
        emailSubmission: submitted.emailSubmission };
    assert.ok(policy.recoverSubmissionProofs(storedProof, proofTask, 'TARGET@outlook.com', 'client', 2000));
    for (const [field, value] of [['runId', 'other'], ['sessionId', 'other'], ['bindingId', 'other'], ['clientId', 'other']] as const) {
        assert.equal(policy.recoverSubmissionProofs({ ...storedProof, [field]: value }, proofTask,
            'target@outlook.com', 'client', 2000), null, `shared proof must bind ${field}`);
    }
    const control = { state: 'running', runId: 'r', sessionId: ID, bindingId: 'b', heartbeat: 1000, manual: false };
    const tab = { runId: 'r', sessionId: ID, bindingId: 'b', expiresAt: 50000 };
    assert.equal(policy.mayAct(control, tab, 2000), true);
    assert.equal(policy.mayAct({ ...control, manual: true }, tab, 2000), false);
    assert.equal(policy.mayAct(control, tab, 22000), false);
    for (const heartbeat of [Number.NaN, Number.POSITIVE_INFINITY, 3000]) {
        assert.equal(policy.mayAct({ ...control, heartbeat }, tab, 2000), false);
    }
    assert.equal(policy.mayAct(control, { ...tab, sessionId: 'other' }, 2000), false);
    assert.equal(policy.mayAct(control, { ...tab, bindingId: 'other' }, 2000), false);
    assert.equal(policy.devicePath('login.microsoftonline.com', '/common/oauth2/deviceauth'), true);
    assert.equal(policy.devicePath('www.microsoft.com', '/link'), true);
    assert.equal(policy.devicePath('evil.example', '/common/oauth2/deviceauth'), false);
    assert.equal(policy.verificationTarget({ verificationUri: 'https://www.microsoft.com/link' }, 'gongxi-helper=test'),
        'https://www.microsoft.com/link#gongxi-helper=test');
    assert.equal(policy.verificationTarget({ verificationUri: 'https://evil.example/link' }, 'gongxi-helper=test'), null);
    const bootstrap = policy.bootstrapTarget({ sessionId: ID, verificationUri: 'https://www.microsoft.com/link' },
        'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb', 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa');
    assert.ok(bootstrap);
    const bootstrapUrl = new URL(bootstrap);
    assert.equal(bootstrapUrl.origin, 'https://outlook.wujiaqiao.dpdns.org');
    assert.equal(bootstrapUrl.pathname, '/reauthorizations');
    const rawBootstrap = bootstrapUrl.hash.slice('#gongxi-helper-launch='.length);
    assert.match(rawBootstrap, /^[A-Za-z0-9_-]+$/);
    const decodedBootstrap = JSON.parse(Buffer.from(rawBootstrap, 'base64url').toString('utf8'));
    assert.deepEqual(decodedBootstrap, { bindingId: 'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb',
        runId: 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa', sessionId: ID, target: 'https://www.microsoft.com/link' });
    assert.equal(policy.bootstrapTarget({ sessionId: ID, verificationUri: 'https://evil.example/link' },
        'bbbbbbbb-bbbb-4bbb-8bbb-bbbbbbbbbbbb', 'aaaaaaaa-aaaa-4aaa-8aaa-aaaaaaaaaaaa'), null);
    assert.equal(policy.workerRequiresPause({ ...control, handoffAt: 1000 }, {}, 22000), true);
    assert.equal(policy.workerRequiresPause({ ...control, handoffAt: 1000 }, { ...tab, heartbeat: 2000, paused: true }, 2000), true);
    assert.equal(policy.workerRequiresPause({ ...control, handoffAt: 1000, manual: true }, { ...tab, heartbeat: 2000, paused: true }, 2000), false);
    const handedOff = { ...control, handoffAt: 1000 };
    assert.equal(policy.workerConfirmedFreshLogin(handedOff,
        { ...tab, heartbeat: 100, paused: false, freshLogin: true, freshLoginAt: 1500 }, 30000), true);
    assert.equal(policy.workerConfirmedFreshLogin(handedOff,
        { ...tab, heartbeat: 2000, paused: false, freshLogin: true, freshLoginAt: 500 }, 2000), false);
    for (const freshLoginAt of [Number.NaN, Number.POSITIVE_INFINITY, 3000]) {
        assert.equal(policy.workerConfirmedFreshLogin(handedOff,
            { ...tab, heartbeat: 2000, paused: false, freshLogin: true, freshLoginAt }, 2000), false);
    }
    assert.equal(policy.workerConfirmedFreshLogin(handedOff,
        { ...tab, bindingId: 'tampered', heartbeat: 2000, paused: false, freshLogin: true, freshLoginAt: 1500 }, 2000), false);
    assert.equal(policy.workerRequiresPause(handedOff, { ...tab, heartbeat: Number.NaN }, 2000), true);
    assert.equal(policy.workerRequiresPause(handedOff, { ...tab, heartbeat: 3000 }, 2000), true);
    assert.equal(policy.workerConfirmedFreshLogin(handedOff,
        { ...tab, heartbeat: 2000, paused: false, freshLogin: false, freshLoginAt: 1500 }, 2000), false);
    const manualSuccess = { runId: 'r', sessionId: ID, bindingId: 'b', confirmedAt: 2500 };
    assert.equal(policy.manualSuccessConfirmed({ ...handedOff, manualSuccess }, { sessionId: ID, status: 'SUCCEEDED' }, 3000), true);
    assert.equal(policy.manualSuccessConfirmed({ ...handedOff, manualSuccess }, { sessionId: ID, status: 'PENDING' }, 3000), false);
    assert.equal(policy.manualSuccessConfirmed({ ...handedOff, manualSuccess: { ...manualSuccess, confirmedAt: 500 } },
        { sessionId: ID, status: 'SUCCEEDED' }, 3000), false);
    assert.doesNotMatch(source, /@grant\s+none|unsafeWindow|@require\s|postMessage\(/);
    assert.deepEqual([...source.matchAll(/@connect\s+([^\s]+)/g)].map((match) => match[1]), ['outlook.wujiaqiao.dpdns.org']);
    assert.doesNotMatch(source, /GM_openInTab\(`https:\/\/microsoft\.com\/devicelogin/);
    assert.match(source, /cooldownUntil = Date\.now\(\) \+ 10000/);
    assert.match(source, /@sandbox\s+DOM/);
});
