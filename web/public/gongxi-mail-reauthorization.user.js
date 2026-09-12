// ==UserScript==
// @name         GongXi Mail 串行重新授权助手
// @namespace    https://outlook.wujiaqiao.dpdns.org/
// @version      1.0.2
// @description  仅为管理员当前设备授权会话填写登录信息；挑战、未知页面或身份不符时暂停。
// @match        https://outlook.wujiaqiao.dpdns.org/*
// @match        https://microsoft.com/devicelogin*
// @match        https://www.microsoft.com/devicelogin*
// @match        https://microsoft.com/link*
// @match        https://www.microsoft.com/link*
// @match        https://login.microsoftonline.com/*
// @match        https://login.live.com/*
// @connect      outlook.wujiaqiao.dpdns.org
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        GM_getTab
// @grant        GM_saveTab
// @grant        GM_openInTab
// @grant        GM_xmlhttpRequest
// @sandbox      DOM
// @run-at       document-idle
// @noframes
// ==/UserScript==

(function () {
    'use strict';
    const ORIGIN = 'https://outlook.wujiaqiao.dpdns.org';
    const CONTROL = 'gongxi-helper-control-v1';
    const LEGACY_TICKET = 'gongxi-helper-ticket-v1';
    const TICKET = 'gongxi-helper-ticket-v2:';
    const BINDING = 'gongxi-helper-binding-v2:';
    const SUBMISSION = 'gongxi-helper-submission-v1';
    const HEARTBEAT = 'gongxi-helper-heartbeat-v1';
    const WORKER = 'gongxi-helper-worker-v1';
    const MS_HOSTS = ['microsoft.com', 'www.microsoft.com', 'login.microsoftonline.com', 'login.live.com'];
    // Built-in Microsoft OAuth application ID from Thunderbird's OAuth2Providers.sys.mjs.
    const THUNDERBIRD_CLIENT_ID = '9e5f94bc-e8a4-4e73-b8be-63364c29d753';
    const ACTIVE = ['STARTING', 'PENDING', 'POLLING'];
    const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
    const normalize = (value) => String(value || '').trim().toLowerCase();
    const otherMicrosoftAccountText = /^(?:Sign in with another Microsoft account|通过其他 Microsoft 帐户登录)$/i;
    const otherAccountText = /^(?:Use another account|Use a different account|使用其他(?:帐户|账户|账号)|使用另一个(?:帐户|账户|账号)|使用其他帐户登录)$/i;
    const chooseOtherAction = (value) => otherMicrosoftAccountText.test(value) ? 'choose-other-microsoft' :
        otherAccountText.test(value) ? 'choose-other' : null;
    const secretShape = (value) => typeof value === 'string' && /^[A-Za-z0-9_-]{43}$/.test(value);
    const uuidShape = (value) => typeof value === 'string' && /^[a-f0-9]{8}-[a-f0-9]{4}-[1-8][a-f0-9]{3}-[89ab][a-f0-9]{3}-[a-f0-9]{12}$/i.test(value);
    const devicePath = (host, path) =>
        ['microsoft.com', 'www.microsoft.com'].includes(host) && /^\/(?:devicelogin|link)\/?$/i.test(path) ||
        host === 'login.microsoftonline.com' && /^\/(?:common|consumers)\/oauth2\/deviceauth\/?$/i.test(path) ||
        host === 'login.live.com' && path === '/oauth20_remoteconnect.srf';

    function verificationTarget(session, fragment) {
        const raw = session?.verificationUri || session?.verificationUriComplete;
        try {
            const target = new URL(raw);
            if (target.protocol !== 'https:' || !MS_HOSTS.includes(target.hostname) ||
                target.username || target.password || !devicePath(target.hostname, target.pathname)) return null;
            target.hash = fragment;
            return target.href;
        } catch { return null; }
    }

    function bootstrapTarget(session, bindingId, runId) {
        const verification = verificationTarget(session, '');
        if (!verification || !uuidShape(bindingId) || !uuidShape(runId) || !uuidShape(session?.sessionId)) return null;
        const payload = btoa(JSON.stringify({ bindingId, runId, sessionId: session.sessionId, target: verification }))
            .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');
        return `${ORIGIN}/reauthorizations#gongxi-helper-launch=${payload}`;
    }

    function submittedDeviceCode(progress, expectedClientId) {
        const proof = progress?.deviceSubmission;
        return !!proof && proof.runId === progress.runId && proof.sessionId === progress.sessionId &&
            proof.bindingId === progress.bindingId && typeof expectedClientId === 'string' && expectedClientId.length > 0 &&
            proof.clientId === expectedClientId && Number.isFinite(proof.submittedAt);
    }
    function submittedAccountEmail(progress, expectedEmail, expectedClientId) {
        const proof = progress?.emailSubmission;
        return submittedDeviceCode(progress, expectedClientId) && !!proof && proof.runId === progress.runId &&
            proof.sessionId === progress.sessionId && proof.bindingId === progress.bindingId &&
            proof.clientId === expectedClientId && normalize(proof.email) === normalize(expectedEmail) &&
            Number.isFinite(proof.submittedAt) && proof.submittedAt >= progress.deviceSubmission.submittedAt;
    }
    function submittedPassword(progress, expectedEmail, expectedClientId) {
        const proof = progress?.passwordSubmission;
        return submittedAccountEmail(progress, expectedEmail, expectedClientId) && progress.passwordSent === true &&
            progress.actions?.password === true && !!proof && proof.runId === progress.runId &&
            proof.sessionId === progress.sessionId && proof.bindingId === progress.bindingId &&
            proof.clientId === expectedClientId && normalize(proof.email) === normalize(expectedEmail) &&
            Number.isFinite(proof.submittedAt) && proof.submittedAt >= progress.emailSubmission.submittedAt;
    }
    function submissionTimelineValid(progress, now) {
        const timestamps = [progress?.boundAt, progress?.deviceSubmission?.submittedAt,
            progress?.emailSubmission?.submittedAt, progress?.passwordSubmission?.submittedAt];
        return timestamps.every(Number.isFinite) && timestamps.every((value, index) => index === 0 || value >= timestamps[index - 1]) &&
            timestamps[timestamps.length - 1] <= now && Number.isFinite(progress?.expiresAt) && now < progress.expiresAt;
    }
    function recoverSubmissionProofs(stored, progress, expectedEmail, expectedClientId, now) {
        if (!stored || stored.version !== 1 || stored.runId !== progress?.runId ||
            stored.sessionId !== progress?.sessionId || stored.bindingId !== progress?.bindingId ||
            stored.clientId !== expectedClientId || normalize(stored.email) !== normalize(expectedEmail) ||
            !Number.isFinite(progress?.boundAt) || stored.boundAt !== progress.boundAt ||
            !Number.isFinite(progress?.expiresAt) || stored.expiresAt !== progress.expiresAt ||
            now < progress.boundAt || now >= progress.expiresAt || progress.actions?.device !== true) return null;
        const deviceSubmission = stored.deviceSubmission;
        const candidate = { ...progress, deviceSubmission };
        if (!submittedDeviceCode(candidate, expectedClientId) || deviceSubmission.submittedAt < progress.boundAt ||
            deviceSubmission.submittedAt > now || deviceSubmission.submittedAt >= progress.expiresAt) return null;
        if (stored.emailSubmission === undefined) return { deviceSubmission };
        candidate.emailSubmission = stored.emailSubmission;
        if (progress.actions?.email !== true || !submittedAccountEmail(candidate, expectedEmail, expectedClientId) ||
            stored.emailSubmission.submittedAt > now || stored.emailSubmission.submittedAt >= progress.expiresAt) return null;
        if (stored.passwordSubmission === undefined) return { deviceSubmission, emailSubmission: stored.emailSubmission };
        candidate.passwordSubmission = stored.passwordSubmission;
        if (!submittedPassword(candidate, expectedEmail, expectedClientId) ||
            stored.passwordSubmission.submittedAt > now || stored.passwordSubmission.submittedAt >= progress.expiresAt) return null;
        return { deviceSubmission, emailSubmission: stored.emailSubmission, passwordSubmission: stored.passwordSubmission };
    }
    function clientCompatible(view, expectedClientId) {
        return typeof expectedClientId === 'string' && expectedClientId.length > 0 &&
            Array.isArray(view.clientIds) &&
            view.clientIds.every((clientId) => clientId === expectedClientId);
    }
    function scopesCompatible(view) {
        if (view.scopeDeclared === true && (!Array.isArray(view.scopes) || view.scopes.length === 0)) return false;
        if (!Array.isArray(view.scopes) || view.scopes.length === 0) return true;
        const normalized = view.scopes.map((scope) => normalize(scope).replace(/^https:\/\/graph\.microsoft\.com\//, ''));
        const allowed = new Set(['openid', 'profile', 'offline_access', 'user.read', 'mail.readwrite']);
        return normalized.includes('mail.readwrite') && normalized.every((scope) => allowed.has(scope));
    }
    // Pure, intentionally conservative policy. Unknown layouts and post-device steps without
    // a tab-private submission proof never submit. Microsoft normally carries the OAuth client
    // in server-side flow state after device-code entry, so client_id is checked when present but
    // is not required to be repeated in every later page URL or form action.
    function decide(view, expectedEmail, expectedClientId, progress) {
        if (view.error) return 'pause';
        // A password entry affordance can also be rendered on recovery, CAPTCHA and MFA pages.
        // Those hard challenges always win over the one supported passwordless-first layout.
        if (view.hardChallenge || view.challenge && view.passwordAlternative !== true) return 'pause';
        // Never interpret a mixed terminal/login transition as a password page, and never
        // automate Microsoft's generic continue affordance.
        if (view.continue || view.consent && (view.device || view.email || view.password || view.passwordOption || view.picker)) return 'pause';
        // Only Microsoft's specifically identified "Verify your email" proof chooser is allowed
        // to switch to the explicit password alternative.
        if (view.passwordOption) {
            if (view.identities.some((identity) => normalize(identity) !== normalize(expectedEmail))) return 'mismatch';
            return view.passwordAlternative === true && submittedAccountEmail(progress, expectedEmail, expectedClientId) &&
                clientCompatible(view, expectedClientId) ? 'choose-password' : 'pause';
        }
        if (view.consent) {
            const identityMatched = view.identities.length === 1 &&
                normalize(view.identities[0]) === normalize(expectedEmail);
            // The official client ID plus this tab's device-code proof anchors the page to the
            // server's fixed DEVICE_SCOPES request. The explicit UI checks are defense in depth.
            return expectedClientId === THUNDERBIRD_CLIENT_ID && identityMatched &&
                submittedPassword(progress, expectedEmail, expectedClientId) &&
                submissionTimelineValid(progress, Date.now()) &&
                clientCompatible(view, expectedClientId) && scopesCompatible(view) &&
                view.applicationThunderbird === true && view.permissionsExpected === true &&
                view.permissionsEnumerated === true && view.permissionsUnexpected !== true &&
                !!view.consentAccept && view.consentAcceptCount === 1
                ? 'accept-consent' : 'pause';
        }
        if (view.device && view.devicePath) return typeof expectedClientId === 'string' && expectedClientId.length > 0 ? 'device' : 'pause';
        const submitted = submittedDeviceCode(progress, expectedClientId);
        const compatible = clientCompatible(view, expectedClientId);
        // Account tiles intentionally contain identities other than the queued account. Always
        // choose the explicit "use another account" tile instead of accepting a cached session.
        if (view.picker) return submitted && compatible &&
            ['choose-other-microsoft', 'choose-other'].includes(view.pickerAction) ? view.pickerAction : 'pause';
        if (view.email) return submitted && compatible ? 'email' : 'pause';
        if (view.identities.some((identity) => normalize(identity) !== normalize(expectedEmail))) return 'mismatch';
        const matched = view.identities.some((identity) => normalize(identity) === normalize(expectedEmail));
        // A matching cached-account password page is still not enough: this helper must have
        // submitted the queued email itself during this bound device flow first.
        if (view.password) return submittedAccountEmail(progress, expectedEmail, expectedClientId) && compatible && matched ? 'password' : 'pause';
        if (view.staySignedIn && matched) return submittedPassword(progress, expectedEmail, expectedClientId) &&
            submissionTimelineValid(progress, Date.now()) && compatible ? 'stay-no' : 'pause';
        if (view.done) return 'wait';
        return 'unknown';
    }
    function mayAct(control, tab, now) {
        return !!control && control.state === 'running' && !control.manual && control.runId === tab.runId &&
            control.sessionId === tab.sessionId && control.bindingId === tab.bindingId &&
            recentTimestamp(control.heartbeat, now, 20000) && now < tab.expiresAt;
    }
    function nextDelay(session, now) {
        return Math.max(1000, Date.parse(session.nextPollAt) - Date.parse(session.serverTime || new Date(now).toISOString()));
    }
    function recentTimestamp(value, now, maxAge) {
        return Number.isFinite(value) && value <= now && now - value < maxAge;
    }
    function workerRequiresPause(control, worker, now) {
        if (!control?.sessionId || control.manual) return false;
        if (!Number.isFinite(control.handoffAt) || control.handoffAt > now) return true;
        return worker?.runId === control.runId && worker.sessionId === control.sessionId && worker.bindingId === control.bindingId
            ? !!worker.paused || !recentTimestamp(worker.heartbeat, now, 20000)
            : !recentTimestamp(control.handoffAt, now, 20000);
    }
    function workerConfirmedFreshLogin(control, worker, now) {
        return !!control?.sessionId && worker?.runId === control.runId && worker.sessionId === control.sessionId &&
            worker.bindingId === control.bindingId && worker.freshLogin === true && !worker.paused &&
            Number.isFinite(control.handoffAt) && control.handoffAt <= now &&
            Number.isFinite(worker.freshLoginAt) && worker.freshLoginAt >= control.handoffAt && worker.freshLoginAt <= now;
    }
    function manualSuccessConfirmed(control, session, now) {
        const proof = control?.manualSuccess;
        return session?.status === 'SUCCEEDED' && proof?.runId === control.runId && proof.sessionId === control.sessionId &&
            proof.sessionId === session.sessionId && proof.bindingId === control.bindingId &&
            Number.isFinite(control.handoffAt) && control.handoffAt <= now &&
            Number.isFinite(proof.confirmedAt) && proof.confirmedAt >= control.handoffAt && proof.confirmedAt <= now;
    }
    function mergeSessionSnapshot(previous, incoming) {
        if (!previous || previous.sessionId !== incoming?.sessionId) return previous;
        return previous.status === 'SUCCEEDED' && incoming.status !== 'SUCCEEDED' ? previous : incoming;
    }

    // The test harness evaluates these pure functions in an isolated Node VM.
    if (typeof document === 'undefined') {
        if (typeof module !== 'undefined') module.exports = { decide, mayAct, devicePath, nextDelay, secretShape,
            submittedDeviceCode, submittedAccountEmail, submittedPassword, submissionTimelineValid,
            recoverSubmissionProofs, verificationTarget, bootstrapTarget,
            workerRequiresPause, workerConfirmedFreshLogin, manualSuccessConfirmed };
        return;
    }
    if (window.top !== window.self || location.protocol !== 'https:') return;
    function readControl() {
        const value = GM_getValue(CONTROL, null);
        const beat = GM_getValue(HEARTBEAT, null);
        return value ? { ...value, heartbeat: beat?.runId === value.runId ? beat.heartbeat : 0 } : null;
    }

    function panel(title) {
        const host = document.createElement('div');
        host.id = 'gongxi-helper-panel';
        host.style.cssText = 'position:fixed;right:18px;bottom:18px;z-index:2147483647;max-width:360px';
        const root = host.attachShadow({ mode: 'closed' });
        const box = document.createElement('div');
        box.style.cssText = 'background:#fff;color:#172554;border:2px solid #2563eb;border-radius:10px;padding:16px;box-shadow:0 5px 24px #0003;font:14px/1.6 system-ui';
        const heading = document.createElement('strong');
        heading.textContent = title;
        const status = document.createElement('div');
        status.style.cssText = 'margin:8px 0;white-space:pre-wrap;overflow-wrap:anywhere';
        box.append(heading, status);
        root.append(box);
        document.body.append(host);
        return {
            set: (value) => { status.textContent = value; },
            button: (label, action) => {
                const button = document.createElement('button');
                button.textContent = label;
                button.style.cssText = 'margin:4px 8px 0 0;padding:5px 10px;cursor:pointer';
                button.addEventListener('click', () => { void action(); });
                box.append(button);
            },
        };
    }

    async function adminMain() {
        const ui = panel('GongXi Mail 授权助手');
        let running = false;
        let quitting = false;
        let tab = null;
        let current = null;
        let runId = null;
        let bindingId = null;
        let beat = null;
        let cooldownUntil = 0;
        let generation = 0;
        let continuePending = false;
        ui.set('先用 1–2 个授权账号试运行。未知页面自动暂停。');
        async function request(path, body) {
            // This function exists only on the backend origin. Never GM-store JWT.
            if (location.origin !== ORIGIN || !path.startsWith('/admin/email-reauthorizations/')) throw new Error('origin');
            const jwt = localStorage.getItem('token');
            if (!jwt) throw new Error('login');
            const response = await fetch(ORIGIN + path, {
                method: body === undefined ? 'GET' : 'POST', cache: 'no-store', redirect: 'error', credentials: 'omit',
                headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${jwt}` },
                ...(body === undefined ? {} : { body: JSON.stringify(body) }),
            });
            const value = await response.json();
            if (!response.ok || value.success !== true) throw new Error('request');
            return value.data;
        }
        function control() { return readControl(); }
        function update(fields) {
            const value = control();
            if (value?.runId === runId) GM_setValue(CONTROL, { ...value, ...fields });
        }
        function captureCurrentContext() {
            const state = control();
            const sessionId = current?.sessionId ?? null;
            const activeBindingId = bindingId ?? null;
            if (!runId || !state || state.runId !== runId || state.sessionId !== sessionId ||
                (state.bindingId ?? null) !== activeBindingId || state.generation !== generation) return null;
            return { runId, sessionId, bindingId: activeBindingId, generation };
        }
        function contextIsCurrent(snapshot) {
            if (!snapshot || runId !== snapshot.runId || generation !== snapshot.generation ||
                (current?.sessionId ?? null) !== snapshot.sessionId || (bindingId ?? null) !== snapshot.bindingId) return false;
            const state = control();
            return state?.runId === snapshot.runId && state.sessionId === snapshot.sessionId &&
                (state.bindingId ?? null) === snapshot.bindingId && state.generation === snapshot.generation;
        }
        function applySessionSnapshot(snapshot, incoming) {
            if (!contextIsCurrent(snapshot) || incoming?.sessionId !== snapshot.sessionId) return false;
            current = mergeSessionSnapshot(current, incoming);
            return true;
        }
        function clearHandoff(id = bindingId) {
            if (!id) return;
            GM_deleteValue(`${BINDING}${id}`);
            GM_deleteValue(`${TICKET}${id}`);
        }
        function pause(reason = '管理员暂停；当前 Microsoft 页面请人工完成') {
            clearHandoff();
            generation += 1;
            update({ state: 'paused', generation, manual: !!current, reason });
        }
        function checkWorker() {
            const state = control();
            const now = Date.now();
            if (manualSuccessConfirmed(state, current, now)) return;
            const worker = GM_getValue(WORKER, null);
            if (workerRequiresPause(state, worker, now)) {
                pause(worker?.paused ? worker.reason : 'Microsoft 助手失联或进入未授权域名，已转人工');
            }
        }
        async function finish() {
            if (!runId || control()?.runId !== runId) { ui.set('此后台标签页没有运行助手。'); return; }
            quitting = true;
            generation += 1;
            update({ state: 'stopped', generation });
            clearHandoff();
            GM_deleteValue(LEGACY_TICKET);
            document.documentElement.removeAttribute('data-gongxi-helper-busy');
            ui.set('助手已结束。刷新后台队列可人工继续当前会话。');
        }
        ui.button('启动 / 继续队列', async () => {
            if (location.pathname !== '/reauthorizations') {
                ui.set('请在后台「批量重新授权」页面启动。'); return;
            }
            if (running) {
                if (continuePending) return;
                continuePending = true;
                try {
                    const state = control();
                    if (current && state?.manual) {
                        try {
                            const readContext = captureCurrentContext();
                            if (!readContext) return;
                            let observed = await request(`/admin/email-reauthorizations/${readContext.sessionId}`);
                            if (!applySessionSnapshot(readContext, observed)) return;
                            if (current.status === 'PENDING' && Date.parse(current.nextPollAt) <= Date.parse(current.serverTime)) {
                                const pollContext = captureCurrentContext();
                                if (!pollContext) return;
                                observed = await request(`/admin/email-reauthorizations/${pollContext.sessionId}/poll`, {});
                                if (!applySessionSnapshot(pollContext, observed)) return;
                            }
                        } catch {
                            ui.set('无法核验当前会话，请稍后重试。'); return;
                        }
                        if (current.status !== 'SUCCEEDED') {
                            ui.set('当前项尚未通过后台核验，请先在 Microsoft 页面完成人工验证。'); return;
                        }
                        const confirmedControl = control();
                        const now = Date.now();
                        if (!confirmedControl || confirmedControl.runId !== runId ||
                            confirmedControl.sessionId !== current.sessionId || confirmedControl.bindingId !== bindingId ||
                            confirmedControl.generation !== generation) return;
                        update({ state: 'running', manual: false, reason: '', manualSuccess: { runId,
                            sessionId: current.sessionId, bindingId: confirmedControl.bindingId, confirmedAt: now } });
                    } else update({ state: 'running', manual: false, reason: '' });
                } finally {
                    continuePending = false;
                }
                return;
            }
            if (!navigator.locks) { ui.set('浏览器不支持单队列互斥锁，请使用最新版 Chrome / Edge。'); return; }
            running = true;
            quitting = false;
            try {
                await navigator.locks.request('gongxi-reauthorization-helper', { ifAvailable: true }, async (lock) => {
                    if (!lock) { ui.set('另一个后台标签页正在运行助手。'); return; }
                    runId = crypto.randomUUID();
                    GM_deleteValue(LEGACY_TICKET);
                    generation = 0;
                    GM_setValue(CONTROL, { runId, state: 'running', sessionId: null, bindingId: null,
                        generation, manual: false, heartbeat: Date.now() });
                    GM_setValue(HEARTBEAT, { runId, heartbeat: Date.now() });
                    document.documentElement.setAttribute('data-gongxi-helper-busy', 'true');
                    beat = setInterval(() => {
                        if (location.pathname !== '/reauthorizations') pause();
                        GM_setValue(HEARTBEAT, { runId, heartbeat: Date.now() });
                    }, 2000);
                    while (!quitting) {
                        checkWorker();
                        const state = control();
                        if (!state || state.runId !== runId || state.state === 'stopped') break;
                        if (!current) {
                            if (state.state !== 'running') { ui.set('队列已暂停。点击继续队列后才会开始下一项。'); await sleep(1000); continue; }
                            if (Date.now() < cooldownUntil) {
                                ui.set(`当前项已成功，${Math.ceil((cooldownUntil - Date.now()) / 1000)} 秒后开始下一项。`);
                                await sleep(1000); continue;
                            }
                            const candidateContext = captureCurrentContext();
                            if (!candidateContext) continue;
                            const candidates = await request('/admin/email-reauthorizations/candidates');
                            if (quitting || !contextIsCurrent(candidateContext) || control()?.state !== 'running') continue;
                            const candidate = candidates.find((value) => value.supported);
                            if (!candidate) { ui.set('队列已完成，没有待重新授权的受支持邮箱。'); break; }
                            let selected = candidate.activeSession;
                            if (!selected) {
                                const startContext = captureCurrentContext();
                                if (!startContext) continue;
                                selected = await request('/admin/email-reauthorizations/start', { emailId: candidate.emailId });
                                if (quitting || !contextIsCurrent(startContext) || control()?.state !== 'running') continue;
                            }
                            generation += 1;
                            current = selected;
                            bindingId = null;
                            update({ sessionId: current.sessionId, bindingId: null, generation,
                                manual: false, handoffAt: null, manualSuccess: null });
                            if (quitting || control()?.state !== 'running') { update({ manual: true }); continue; }
                            if (!['PENDING', 'POLLING'].includes(current.status)) {
                                pause('会话未就绪，请人工检查'); continue;
                            }
                            // Bind the exact child tab before a ticket exists in shared GM storage.
                            clearHandoff();
                            bindingId = crypto.randomUUID();
                            const openedAt = Date.now();
                            generation += 1;
                            update({ bindingId, generation, handoffAt: openedAt });
                            if (tab) { tab.close(); tab = null; }
                            const bootstrap = bootstrapTarget(current, bindingId, runId);
                            if (!bootstrap) {
                                pause('Microsoft 未返回受支持的设备授权地址，已转人工');
                                continue;
                            }
                            tab = GM_openInTab(bootstrap, { active: true, insert: true, setParent: true });
                            let bound = null;
                            const bindDeadline = openedAt + 15000;
                            while (!quitting && Date.now() < bindDeadline && control()?.state === 'running' && !tab?.closed) {
                                const candidate = GM_getValue(`${BINDING}${bindingId}`, null);
                                if (candidate?.runId === runId && candidate.sessionId === current.sessionId &&
                                    candidate.bindingId === bindingId && candidate.boundAt >= openedAt && candidate.boundAt <= Date.now()) {
                                    bound = candidate; break;
                                }
                                await sleep(200);
                            }
                            if (!bound) {
                                clearHandoff();
                                pause('无法安全绑定新 Microsoft 标签页，已转人工');
                                continue;
                            }
                            // Exactly one handoff per session. A lost/expired ticket requires manual recovery.
                            const ticketContext = captureCurrentContext();
                            if (!ticketContext) { update({ manual: true }); continue; }
                            const issued = await request(`/admin/email-reauthorizations/${ticketContext.sessionId}/helper-ticket`, {});
                            if (quitting || !contextIsCurrent(ticketContext) || control()?.state !== 'running') {
                                issued.ticket = ''; clearHandoff(); update({ manual: true }); continue;
                            }
                            GM_setValue(`${TICKET}${bindingId}`, { ticket: issued.ticket, runId, sessionId: current.sessionId,
                                bindingId, boundAt: bound.boundAt, publishedAt: Date.now(), expiresAt: Date.parse(issued.expiresAt) });
                            GM_deleteValue(`${BINDING}${bindingId}`);
                            issued.ticket = '';
                        }
                        const latest = control();
                        ui.set(`${current.email}\n${latest?.state === 'paused' || latest?.manual ? (latest.reason || '当前项已转人工；成功后可继续下一项') : '正在等待 Microsoft 授权并核验身份…'}`);
                        if (tab?.closed && latest?.state === 'running' &&
                            !manualSuccessConfirmed(latest, current, Date.now())) {
                            pause('Microsoft 标签页已关闭，请人工检查当前项');
                        }
                        await sleep(Math.min(5000, nextDelay(current, Date.now())));
                        if (quitting) break;
                        // Only the backend JWT holder polls or advances the queue.
                        const readContext = captureCurrentContext();
                        if (!readContext) continue;
                        let observed = await request(`/admin/email-reauthorizations/${readContext.sessionId}`);
                        if (!applySessionSnapshot(readContext, observed)) continue;
                        if (current.status === 'PENDING' && Date.parse(current.nextPollAt) <= Date.parse(current.serverTime)) {
                            const pollContext = captureCurrentContext();
                            if (!pollContext) continue;
                            observed = await request(`/admin/email-reauthorizations/${pollContext.sessionId}/poll`, {});
                            if (!applySessionSnapshot(pollContext, observed)) continue;
                        }
                        checkWorker();
                        if (current.status === 'SUCCEEDED') {
                            const succeededControl = control();
                            const succeededWorker = GM_getValue(WORKER, null);
                            const succeededAt = Date.now();
                            if (!succeededControl || current.sessionId !== succeededControl.sessionId ||
                                succeededControl.runId !== runId || succeededControl.bindingId !== bindingId ||
                                succeededControl.generation !== generation) continue;
                            if (!workerConfirmedFreshLogin(succeededControl, succeededWorker, succeededAt) &&
                                !manualSuccessConfirmed(succeededControl, current, succeededAt)) {
                                pause('未确认当前账号经过使用其他账号/全新邮箱及密码提交，已停止队列');
                                continue;
                            }
                            cooldownUntil = Date.now() + 10000;
                            clearHandoff();
                            generation += 1;
                            bindingId = null;
                            if (tab) { tab.close(); tab = null; }
                            current = null;
                            update({ sessionId: null, bindingId: null, generation,
                                manual: false, handoffAt: null, manualSuccess: null });
                        } else if (!ACTIVE.includes(current.status)) {
                            pause('会话失败、取消、过期或身份不匹配。请结束助手，在后台检查 / 重试。');
                        }
                    }
                });
            } catch {
                generation += 1;
                update({ state: 'stopped', generation, manual: true });
                ui.set('助手已停止：后台请求或安全凭据失败。请结束助手并刷新队列人工检查；同一会话不会重复领密码。');
            } finally {
                if (beat) clearInterval(beat);
                if (runId && control()?.runId === runId) {
                    generation += 1;
                    update({ state: 'stopped', generation });
                    clearHandoff();
                    GM_deleteValue(LEGACY_TICKET);
                    document.documentElement.removeAttribute('data-gongxi-helper-busy');
                }
                running = false;
                generation += 1;
                current = null;
                bindingId = null;
                runId = null;
            }
        });
        ui.button('暂停（当前项转人工）', async () => pause());
        ui.button('结束助手 / 切换人工', finish);
        window.addEventListener('pagehide', () => {
            if (runId && control()?.runId === runId) {
                generation += 1;
                update({ state: 'stopped', generation }); clearHandoff(); GM_deleteValue(LEGACY_TICKET);
            }
        });
    }

    function helperRequest(path, capability, body = {}) {
        return new Promise((resolve, reject) => {
            if (!['claim', 'current', 'password'].includes(path)) { reject(new Error('path')); return; }
            GM_xmlhttpRequest({
                method: 'POST', url: `${ORIGIN}/api/reauthorization-helper/${path}`, anonymous: true,
                redirect: 'error', timeout: 15000, responseType: 'json',
                headers: { 'Content-Type': 'application/json', ...(capability ? { 'X-Reauthorization-Capability': capability } : {}) },
                data: JSON.stringify(body),
                onload: (response) => {
                    if (response.finalUrl !== `${ORIGIN}/api/reauthorization-helper/${path}` || response.status !== 200 || response.response?.success !== true) {
                        reject(new Error('helper')); return;
                    }
                    resolve(response.response.data);
                    response.response = null;
                    response.responseText = '';
                },
                onerror: () => reject(new Error('network')),
                ontimeout: () => reject(new Error('timeout')),
            });
        });
    }
    const visible = (element) => !!element && element.getClientRects().length > 0 && !element.disabled && element.getAttribute('aria-hidden') !== 'true';
    const allVisible = (selector) => [...document.querySelectorAll(selector)].filter(visible);
    const one = (selector) => [...document.querySelectorAll(selector)].find(visible) || null;
    const words = (element) => (element?.innerText || element?.value || element?.textContent || '').trim();
    const elementForm = (element) => element?.form || element?.closest?.('form') || null;
    function effectiveSubmission(submitter) {
        if (!visible(submitter)) return null;
        const form = elementForm(submitter);
        if (!form) return null;
        const actionOverride = submitter.getAttribute?.('formaction');
        const methodOverride = submitter.getAttribute?.('formmethod');
        const rawAction = actionOverride === null || actionOverride === undefined
            ? (form.action || location.href)
            : (submitter.formAction || actionOverride || location.href);
        const rawMethod = methodOverride === null || methodOverride === undefined
            ? (form.method || 'get')
            : (submitter.formMethod || methodOverride || 'get');
        try {
            return { form, action: new URL(rawAction, location.href), method: String(rawMethod).trim().toLowerCase() || 'get' };
        } catch { return null; }
    }
    function safeMicrosoftUrl(target) {
        return target?.protocol === 'https:' && MS_HOSTS.includes(target.hostname) && !target.username && !target.password;
    }
    function urlClientIds(raw) {
        try { return new URL(raw, location.href).searchParams.getAll('client_id'); } catch { return []; }
    }
    function urlScopes(raw) {
        try {
            return new URL(raw, location.href).searchParams.getAll('scope').flatMap((value) => {
                const trimmed = value.trim();
                return trimmed ? trimmed.split(/\s+/) : [''];
            });
        }
        catch { return []; }
    }
    function inputValues(name) {
        return [...document.querySelectorAll(`input[name="${name}"]`)].map((element) =>
            String(element.value ?? element.getAttribute?.('value') ?? '').trim());
    }
    function oauthParameterDeclared(element, name) {
        const submission = effectiveSubmission(element);
        const inUrl = (raw) => {
            try { return new URL(raw, location.href).searchParams.has(name); } catch { return false; }
        };
        return inUrl(location.href) || !!submission && inUrl(submission.action.href) ||
            document.querySelectorAll(`input[name="${name}"]`).length > 0;
    }
    function oauthPageValues(element, name) {
        const submission = effectiveSubmission(element);
        const result = [];
        const fromUrl = name === 'client_id' ? urlClientIds : urlScopes;
        result.push(...fromUrl(location.href));
        if (submission) result.push(...fromUrl(submission.action.href));
        for (const value of inputValues(name)) {
            if (name === 'scope') result.push(...(value ? value.split(/\s+/).filter(Boolean) : ['']));
            else result.push(value);
        }
        return result;
    }
    function permissionFingerprint(value) {
        return String(value || '').toLowerCase().replace(/[\s.,;:!?()\[\]，。；：！？（）【】]+/g, '');
    }
    function textLines(value) {
        return String(value || '').split(/\r?\n/).map((line) => line.trim().replace(/\s+/g, ' ')).filter(Boolean);
    }
    function supportedConsentChromeLine(value) {
        const text = String(value || '').trim().replace(/\s+/g, ' ');
        return /^(?:Microsoft)$/i.test(text) ||
            /^(?:permissions requested by this app|requested permissions|let this app access your info\??|此应用请求的权限|此应用所请求的权限|请求的权限)$/i.test(text) ||
            /^(?:do you want to allow this app to access your (?:info|information)\??(?:\s*\(1 app, 1 total\))?|是否允许此应用访问(?:你|您)的信息\??(?:\s*[（(]1\s*个应用[，,]\s*共\s*1\s*个[）)])?)$/i.test(text) ||
            /^(?:Thunderbird needs your permission to (?:do|perform) the following:?|Thunderbird 需要得到(?:你|您)的许可才能执行以下操作[:：]?)$/i.test(text) ||
            /^(?:show details|hide details|learn more|显示详细信息|隐藏详细信息|了解详细信息)$/i.test(text) ||
            /^(?:accept|reject|cancel|接受|拒绝|取消)$/i.test(text) ||
            /^(?:sign out|terms of (?:use|service)|privacy (?:statement|policy)|help|退出登录|使用条款|服务条款|隐私(?:声明|和 cookie)|帮助)$/i.test(text) ||
            /^(?:accepting these permissions means that you allow this app to use your data as specified in (?:its|their) terms of service and privacy statement\.?|the publisher has not provided links? to (?:its|their) terms(?: of service)? for you to review\.?|you can revoke these permissions at any time in your account\.?)$/i.test(text) ||
            /^(?:接受(?:这些权限|后)?[，,]?(?:即表(?:示|明))?你允许此应用按照(?:其)?服务条款和隐私声明(?:中的相关规定|中的规定)?使用你的(?:数据|信息)。?|发布者(?:尚未|未)提供(?:指向)?(?:其)?条款链接(?:以|供你)?查看。?|你可以(?:随时)?(?:在你的帐户中|从你的帐户中|在 https:\/\/microsoft\.com\/consent )?(?:撤销|更改)这些权限。?)$/i.test(text) ||
            /^(?:如果这不是(?:你|您)的设备[，,]?请使用专用浏览。?)$/i.test(text) ||
            /^https:\/\/(?:myapps\.microsoft\.com|account\.live\.com\/consent\/manage)\/?$/i.test(text);
    }
    function supportedPermission(value) {
        const text = String(value || '').trim().replace(/\s+/g, ' ');
        const mail = /^(?:read and write (?:access to )?your (?:mail|email)(?: thunderbird (?:will be able to|can) read, update, create(?:,| and) delete (?:mail|email)(?: messages?)? in your mailbox\.?)?|(?:你|您)的邮件的读写访问权限(?: thunderbird 将能够读取、更新、创建和删除邮箱中的电子邮件。?(?:不包括发送邮件(?:的)?权限。?)?)?|读取和写入(?:你的|您的)?邮件|读写(?:你的|您的)?邮件|读取、写入(?:你的|您的)?邮件)$/i;
        if (mail.test(text)) return 'mail.readwrite';
        if (/^(?:sign you in and read your profile|登录并读取(?:你的|您的)?个人资料)$/i.test(text)) return 'user.read';
        if (/^(?:maintain access to data you have given it access to|保持对(?:你|您)已授予访问权限的数据的访问权限)$/i.test(text)) return 'offline_access';
        return null;
    }
    function consentPermissionEvidence(pageText, applicationComponents, identities) {
        const containers = allVisible('#permissionsList, #idDiv_ConsentScopes');
        if (containers.length !== 1) return { expected: false, enumerated: false, unexpected: true };
        const container = containers[0];
        const items = [...container.querySelectorAll('li, [role="listitem"]')]
            .filter(visible).map(words).filter(Boolean);
        if (items.length === 0 || permissionFingerprint(words(container)) !== permissionFingerprint(items.join(''))) {
            return { expected: false, enumerated: false, unexpected: true };
        }
        const permissions = items.map(supportedPermission);
        const complete = permissions.every(Boolean) && new Set(permissions).size === permissions.length;
        const expected = complete && permissions.filter((permission) => permission === 'mail.readwrite').length === 1;
        // Prove that the supported permission container covers every permission-like line on the
        // consent page. Known application, identity and action components are removed by exact
        // line equality; all remaining Microsoft chrome must match a small positive allowlist.
        // Thus a second permission region or a loose permission sentence fails closed without a
        // brittle list of dangerous permission keywords.
        const uncovered = textLines(pageText);
        const coveredComponents = [container, ...applicationComponents];
        for (const component of coveredComponents) {
            for (const line of textLines(words(component))) {
                const fingerprint = permissionFingerprint(line);
                const index = uncovered.findIndex((candidate) => permissionFingerprint(candidate) === fingerprint);
                if (index >= 0) uncovered.splice(index, 1);
            }
        }
        // Identity elements are not trusted as a whole: Microsoft layouts may wrap auxiliary
        // text beside the address. Remove only exact email lines and leave every other line for
        // the consent-page allowlist, so permission text cannot hide inside an identity node.
        const identityEmails = identities.flatMap((component) =>
            words(component).match(/[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || []);
        for (const email of identityEmails) {
            const index = uncovered.findIndex((candidate) =>
                permissionFingerprint(candidate) === permissionFingerprint(email));
            if (index >= 0) uncovered.splice(index, 1);
        }
        const pageCovered = uncovered.every(supportedConsentChromeLine);
        return { expected: expected && pageCovered, enumerated: complete && pageCovered,
            unexpected: !complete || !pageCovered };
    }
    function snapshot() {
        const text = document.body.innerText || '';
        const identities = [...document.querySelectorAll('#displayName, #bannerText, #idDiv_PWD_Username, #loginHeader .identity, [data-test-id="user-display-name"], #idDiv_UserTile .table-cell.text-left')]
            .filter(visible).flatMap((node) => words(node).match(/[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || []);
        const email = one('#i0116[name="loginfmt"], #usernameEntry[type="email"][autocomplete~="username"]');
        const password = one('input#i0118[name="passwd"][type="password"], #passwordEntry[type="password"], input[type="password"][autocomplete="current-password"]');
        const device = one('#otc, input[name="user_code"]');
        const actionButtons = allVisible('#idSIButton9, #idSubmit_Consent, #idBtn_Accept, form button[type="submit"], form input[type="submit"], button, input[type="button"], [role="button"]');
        const consentAccepts = actionButtons.filter((node) => /^(?:Accept|接受)$/i.test(words(node)));
        const consentRejects = actionButtons.filter((node) => /^(?:Reject|拒绝)$/i.test(words(node)));
        const submitButtons = actionButtons.filter((node) =>
            !/^(?:Accept|接受|Reject|拒绝)$/i.test(words(node)) &&
            (node.id === 'idSIButton9' || node.matches?.('button[type="submit"], input[type="submit"]')));
        const submit = submitButtons[0] || null;
        const picker = allVisible('#otherTile, #idDiv_UseAnotherAccount, button, a, [role="button"]')
            .find((node) => chooseOtherAction(words(node))) || null;
        const passwordOption = allVisible('button, a, [role="button"]').find((node) =>
            /^(?:Use (?:your )?password|Sign in with (?:your )?password|使用密码|用密码登录|使用密码登录)$/i.test(words(node))) || null;
        const no = one('#idBtn_Back') || allVisible('button, input[type="button"], input[type="submit"]').find((node) =>
            /^(?:No|否)$/i.test(words(node))) || null;
        const isDevicePage = devicePath(location.hostname, location.pathname) && !!device;
        const proofChooser = !!one('#idDiv_SAOTCS_Proofs');
        const verificationEmailHeadings = allVisible('h1, h2, h3, h4, [role="heading"]')
            .filter((node) => /^(?:Verify your email|验证(?:你|您)的电子邮件)$/i.test(words(node)));
        const hardChallengeControl = !!one('#idTxtBx_SAOTCC_OTC, #iOttText, iframe[src*="captcha"], input[name="ProofConfirmation"]');
        const hardChallengeText = /captcha|verify your identity|help us protect|unusual (?:activity|sign.in)|account (?:has been )?locked|security (?:code|info)|authenticator|approve (?:the |a )?(?:sign.in|request)|two.step|recover your account|验证码|验证你的身份|验证您的身份|保护你的帐户|保护您的帐户|帐户已锁定|账户已锁定|异常活动|安全代码|恢复帐户|恢复账户|人机验证|短信验证/i.test(text);
        const oneTimeCodeChallenge = !isDevicePage && !!one('input[name="otc"][autocomplete="one-time-code"]');
        const passwordAlternative = !!passwordOption && proofChooser && verificationEmailHeadings.length === 1 &&
            !hardChallengeControl && !hardChallengeText && !oneTimeCodeChallenge;
        const hardChallenge = hardChallengeControl || hardChallengeText || oneTimeCodeChallenge ||
            proofChooser && !passwordAlternative;
        const challenge = proofChooser || hardChallenge;
        const error = !!one('#passwordError, #usernameError, #errorText, #idDiv_SAOTCC_Error, [role="alert"]') &&
            !!words(one('#passwordError, #usernameError, #errorText, #idDiv_SAOTCC_Error, [role="alert"]')) ||
            /incorrect password|password is incorrect|too many (?:attempts|requests)|try again later|密码不正确|密码错误|尝试次数过多|稍后重试/i.test(text);
        const staySignedIn = /stay signed in|保持登录|保持登入/i.test(text) && visible(no);
        const permissionPage = /permissions requested|let this app access|requested permissions|此应用|应用.*权限|请求的权限/i.test(text);
        const consent = permissionPage || consentAccepts.length > 0 || consentRejects.length > 0;
        const clientTarget = consent ? consentAccepts[0] || consentRejects[0] : email || password ? submit : picker || (staySignedIn ? no : null);
        const explicitApplicationComponents = allVisible('#appDisplayName, #idSpan_AppName, #idDiv_SA_AppName, #idDiv_AppName, [data-test-id="app-display-name"], [data-testid="app-display-name"]');
        const consentHeadings = allVisible('h1, h2, h3, h4, [role="heading"]');
        const supportedConsentHeading = (value) => /^(?:do you want to allow this app to access your (?:info|information)\??(?:\s*\(1 app, 1 total\))?|是否允许此应用访问(?:你|您)的信息\??(?:\s*[（(]1\s*个应用[，,]\s*共\s*1\s*个[）)])?)$/i.test(String(value || '').trim().replace(/\s+/g, ' '));
        // The legacy consumer consent layout renders the verified application name as a link.
        // Accept that narrow fallback only inside the consent form, only when every heading is the
        // known generic consent heading, and only with the official Thunderbird client-ID gate in decide().
        const consentForm = consentAccepts.length === 1 ? elementForm(consentAccepts[0]) : null;
        const legacyApplicationLinks = explicitApplicationComponents.length === 0 && consentForm
            ? allVisible('#idDiv_Consent a, #idDiv_ConsentHeader a, form a').filter((node) =>
                elementForm(node) === consentForm && /^(?:Mozilla Thunderbird|Thunderbird)$/i.test(words(node))) : [];
        const applicationComponents = explicitApplicationComponents.length > 0 ? explicitApplicationComponents :
            consentHeadings.length === 1 && supportedConsentHeading(words(consentHeadings[0])) &&
                legacyApplicationLinks.length === 1 ? legacyApplicationLinks : [];
        const permissionEvidence = consentPermissionEvidence(text, applicationComponents,
            [...document.querySelectorAll('#displayName, #bannerText, #idDiv_PWD_Username, #loginHeader .identity, [data-test-id="user-display-name"], #idDiv_UserTile .table-cell.text-left')].filter(visible));
        return {
            identities, email, password, device, submit, picker,
            pickerAction: picker ? chooseOtherAction(words(picker)) : null,
            passwordOption, passwordAlternative, no, challenge, hardChallenge, error,
            consentAccept: consentAccepts.length === 1 ? consentAccepts[0] : null,
            consentReject: consentRejects.length === 1 ? consentRejects[0] : null,
            consentAcceptCount: consentAccepts.length,
            devicePath: devicePath(location.hostname, location.pathname),
            staySignedIn, clientIds: oauthPageValues(clientTarget, 'client_id'), scopes: oauthPageValues(clientTarget, 'scope'),
            scopeDeclared: oauthParameterDeclared(clientTarget, 'scope'),
            consent, applicationThunderbird: applicationComponents.length === 1 &&
                /^(?:Mozilla Thunderbird|Thunderbird)$/i.test(words(applicationComponents[0])),
            permissionsExpected: permissionEvidence.expected, permissionsEnumerated: permissionEvidence.enumerated,
            permissionsUnexpected: permissionEvidence.unexpected,
            continue: actionButtons.some((node) => /^(?:Continue|继续)$/i.test(words(node))) &&
                /sign(?:ing)? in to|trying to sign in|正在登录|正在登入|尝试登录/i.test(text),
            done: /you have signed in|you('re| are) now signed in|you may (?:now )?close|you can (?:now )?close|已成功登录|现在可以关闭|您已登录|你已登录/i.test(text),
        };
    }
    function compatibleClientIds(rawUrls, expectedClientId) {
        return typeof expectedClientId === 'string' && expectedClientId.length > 0 && rawUrls
            .flatMap((raw) => urlClientIds(raw)).every((value) => value === expectedClientId);
    }
    function safeTarget(element, expectedClientId, requirePost) {
        const submission = effectiveSubmission(element);
        if (!submission || !safeMicrosoftUrl(submission.action) || !safeMicrosoftUrl(new URL(location.href))) return false;
        if (requirePost && submission.method !== 'post') return false;
        return compatibleClientIds([location.href, submission.action.href], expectedClientId);
    }
    function safeEntryTarget(element, expectedClientId) {
        if (!visible(element) || !safeMicrosoftUrl(new URL(location.href))) return null;
        const rawHref = element.getAttribute?.('href');
        if (rawHref !== null && rawHref !== undefined && String(rawHref).trim()) {
            try {
                const target = new URL(element.href || rawHref, location.href);
                return safeMicrosoftUrl(target) && compatibleClientIds([location.href, target.href], expectedClientId)
                    ? { action: target, method: 'get' } : null;
            } catch { return null; }
        }
        const form = elementForm(element);
        if (form) {
            const submission = effectiveSubmission(element);
            return submission && safeMicrosoftUrl(submission.action) &&
                compatibleClientIds([location.href, submission.action.href], expectedClientId) ? submission : null;
        }
        const target = new URL(location.href);
        return compatibleClientIds([target.href], expectedClientId) ? { action: target, method: '' } : null;
    }
    function fill(input, value) {
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
        setter.call(input, value);
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    function actionElements(view, action) {
        const target = action === 'stay-no' ? view.no : ['choose-other-microsoft', 'choose-other'].includes(action) ? view.picker :
            action === 'choose-password' ? view.passwordOption : action === 'accept-consent' ? view.consentAccept : view.submit;
        const input = action === 'device' ? view.device : action === 'email' ? view.email : action === 'password' ? view.password : null;
        return { target, input, form: elementForm(target), inputForm: elementForm(input) };
    }
    function captureAction(action, view, current) {
        const elements = actionElements(view, action);
        if (['choose-other-microsoft', 'choose-other', 'choose-password'].includes(action)) {
            const expectedText = action === 'choose-other-microsoft' ? otherMicrosoftAccountText :
                action === 'choose-other' ? otherAccountText :
                    /^(?:Use (?:your )?password|Sign in with (?:your )?password|使用密码|用密码登录|使用密码登录)$/i;
            const entry = safeEntryTarget(elements.target, current.clientId);
            if (!entry || !expectedText.test(words(elements.target))) return null;
            return { action, href: location.href, target: elements.target, input: null, form: null,
                submissionAction: entry.action.href, submissionMethod: entry.method };
        }
        if (action === 'accept-consent') {
            const submission = effectiveSubmission(elements.target);
            if (!submission || !safeTarget(elements.target, current.clientId, true) ||
                !visible(elements.target) || !/^(?:Accept|接受)$/i.test(words(elements.target))) return null;
            return { action, href: location.href, target: elements.target, input: null, form: elements.form,
                submissionAction: submission.action.href, submissionMethod: submission.method };
        }
        const submission = effectiveSubmission(elements.target);
        if (!submission || !safeTarget(elements.target, current.clientId, ['email', 'password'].includes(action))) return null;
        if (!visible(elements.target) || elements.input && elements.inputForm !== elements.form) return null;
        if (['device', 'email', 'password'].includes(action) && elements.target.id !== 'idSIButton9' &&
            !elements.target.matches?.('button[type="submit"]') &&
            !/^(?:Next|Sign in|Continue|下一步|登录|登入|继续)$/i.test(words(elements.target))) return null;
        return { action, href: location.href, target: elements.target, input: elements.input, form: elements.form,
            submissionAction: submission.action.href, submissionMethod: submission.method };
    }
    function recaptureAction(marker, current, task) {
        if (!marker || task.stopped || !mayAct(readControl(), task, Date.now())) return null;
        const view = snapshot();
        if (decide(view, current.email, current.clientId, task) !== marker.action) return null;
        const fresh = captureAction(marker.action, view, current);
        return fresh && fresh.href === marker.href && fresh.target === marker.target && fresh.input === marker.input &&
            fresh.form === marker.form && fresh.submissionAction === marker.submissionAction &&
            fresh.submissionMethod === marker.submissionMethod ? fresh : null;
    }
    function captureBoundDevicePage() {
        const view = snapshot();
        if (!view.device || !view.devicePath || view.consent || view.continue) return null;
        const elements = actionElements(view, 'device');
        const submission = effectiveSubmission(elements.target);
        if (!submission || !safeMicrosoftUrl(submission.action) || !safeMicrosoftUrl(new URL(location.href)) ||
            elements.inputForm !== elements.form ||
            elements.target.id !== 'idSIButton9' && !/^(?:Next|Continue|下一步|继续)$/i.test(words(elements.target))) return null;
        return { href: location.href, target: elements.target, input: elements.input, form: elements.form,
            submissionAction: submission.action.href, submissionMethod: submission.method };
    }
    function sameBoundDevicePage(marker, task) {
        if (!marker || task.stopped || !mayAct(readControl(), task, Date.now())) return false;
        const fresh = captureBoundDevicePage();
        return !!fresh && fresh.href === marker.href && fresh.target === marker.target && fresh.input === marker.input &&
            fresh.form === marker.form && fresh.submissionAction === marker.submissionAction &&
            fresh.submissionMethod === marker.submissionMethod;
    }
    function readLaunchBinding(hash) {
        const parts = hash.startsWith('#gongxi-helper=') ? hash.slice('#gongxi-helper='.length).split('.') : [];
        return parts.length === 3 && parts.every(uuidShape)
            ? { bindingId: parts[0], runId: parts[1], sessionId: parts[2] }
            : null;
    }
    function readBootstrap(hash) {
        try {
            const prefix = '#gongxi-helper-launch=';
            if (!hash.startsWith(prefix)) return null;
            const raw = hash.slice(prefix.length);
            if (!raw || !/^[A-Za-z0-9_-]+$/.test(raw)) return null;
            const decoded = raw.replace(/-/g, '+').replace(/_/g, '/');
            const payload = JSON.parse(atob(decoded + '='.repeat((4 - decoded.length % 4) % 4)));
            const binding = [payload?.bindingId, payload?.runId, payload?.sessionId].every(uuidShape)
                ? { bindingId: payload.bindingId, runId: payload.runId, sessionId: payload.sessionId } : null;
            const target = verificationTarget({ verificationUri: payload?.target }, '');
            return binding && target ? { ...binding, target } : null;
        } catch { return null; }
    }
    function savedLaunch(value) {
        const target = verificationTarget({ verificationUri: value?.target }, '');
        return value && target && uuidShape(value.bindingId) && uuidShape(value.runId) && uuidShape(value.sessionId) &&
            Number.isFinite(value.expiresAt) && value.expiresAt > Date.now() ? { ...value, target } : null;
    }
    async function bootstrapMain(launch) {
        const target = verificationTarget({ verificationUri: launch?.target }, '');
        if (!target) return;
        const now = Date.now();
        const candidate = { ...launch, target, expiresAt: now + 30000 };
        if (!mayAct(readControl(), candidate, now)) return;
        const tabState = await new Promise((resolve) => GM_getTab(resolve));
        tabState.gongxiHelperLaunch = candidate;
        await new Promise((resolve) => GM_saveTab(tabState, resolve));
        if (!mayAct(readControl(), candidate, Date.now())) return;
        location.replace(target);
    }
    async function microsoftMain() {
        const tabState = await new Promise((resolve) => GM_getTab(resolve));
        let task = tabState.gongxiHelper;
        const launch = readLaunchBinding(location.hash) || savedLaunch(tabState.gongxiHelperLaunch);
        if (task && (!uuidShape(task.runId) || !uuidShape(task.sessionId) || !uuidShape(task.bindingId))) return;
        const save = () => new Promise((resolve) => {
            tabState.gongxiHelper = task;
            GM_saveTab(tabState, resolve);
        });
        const ui = panel('GongXi Mail 当前授权');
        let pendingPasswordInput = null;
        function clearPendingPassword() {
            try {
                if (pendingPasswordInput) fill(pendingPasswordInput, '');
            } catch { /* The page may already have replaced the detached input. */ }
            pendingPasswordInput = null;
        }
        function clearSubmissionProof() {
            const stored = GM_getValue(SUBMISSION, null);
            if (stored?.runId === task.runId && stored.sessionId === task.sessionId && stored.bindingId === task.bindingId) {
                GM_deleteValue(SUBMISSION);
            }
        }
        function restoreSubmissionProofs(current) {
            const restored = recoverSubmissionProofs(GM_getValue(SUBMISSION, null), task,
                current.email, current.clientId, Date.now());
            if (!restored) return;
            task.deviceSubmission = restored.deviceSubmission;
            if (restored.emailSubmission) task.emailSubmission = restored.emailSubmission;
            if (restored.passwordSubmission) task.passwordSubmission = restored.passwordSubmission;
        }
        function persistSubmissionProofs(current) {
            const stored = {
                version: 1, runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                clientId: current.clientId, email: current.email, boundAt: task.boundAt, expiresAt: task.expiresAt,
                deviceSubmission: task.deviceSubmission,
                ...(task.emailSubmission ? { emailSubmission: task.emailSubmission } : {}),
                ...(task.passwordSubmission ? { passwordSubmission: task.passwordSubmission } : {}),
            };
            if (!recoverSubmissionProofs(stored, task, current.email, current.clientId, Date.now())) return false;
            // Legacy GM_setValue is synchronous. This non-secret receipt is durable before the
            // form's default navigation can terminate this document and its GM_saveTab callback.
            GM_setValue(SUBMISSION, stored);
            return true;
        }
        function pause(message) {
            clearPendingPassword();
            GM_deleteValue(`${TICKET}${task.bindingId}`);
            clearSubmissionProof();
            const state = readControl();
            if (state?.runId === task.runId && state.sessionId === task.sessionId && state.bindingId === task.bindingId) {
                GM_setValue(WORKER, { runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                    heartbeat: Date.now(), paused: true, reason: message });
            }
            task.capability = '';
            task.stopped = true;
            void save();
            ui.set(message + '\n请人工处理；后台核验成功前不会进入下一项。');
        }
        let claimPage = captureBoundDevicePage();
        if (!task && launch && !claimPage) {
            const readyDeadline = Math.min(launch.expiresAt ?? Date.now() + 15000, Date.now() + 15000);
            while (!claimPage && Date.now() < readyDeadline && mayAct(readControl(), launch, Date.now()) &&
                safeMicrosoftUrl(new URL(location.href))) {
                await sleep(200);
                claimPage = captureBoundDevicePage();
            }
        }
        if (!task) {
            if (!launch || !claimPage) return;
            const boundAt = Date.now();
            const candidate = { ...launch, boundAt, expiresAt: boundAt + 30000 };
            if (!mayAct(readControl(), candidate, boundAt)) return;
            task = { ...candidate, actions: {} };
            delete tabState.gongxiHelperLaunch;
            await save();
            if (!sameBoundDevicePage(claimPage, task)) { pause('设备授权页在绑定期间发生变化，已暂停'); return; }
            GM_setValue(`${BINDING}${task.bindingId}`, { ...launch, boundAt });
            GM_setValue(WORKER, { ...launch, heartbeat: Date.now(), paused: false });
            if (history?.replaceState) history.replaceState(null, '', location.pathname + location.search);
            claimPage = captureBoundDevicePage();
            if (!claimPage || !mayAct(readControl(), task, Date.now())) { pause('设备授权页在绑定期间发生变化，已暂停'); return; }
        } else if (!task.capability && !claimPage) {
            pause('当前标签页已离开绑定的设备授权页，未领取安全凭据'); return;
        }
        ui.button('暂停并人工处理', async () => pause('管理员已切换人工处理'));
        try {
            if (!task.capability) {
                let handoff = null;
                const ticketKey = `${TICKET}${task.bindingId}`;
                const deadline = Math.min(task.expiresAt, Date.now() + 20000);
                while (mayAct(readControl(), task, Date.now()) && Date.now() < deadline) {
                    if (!sameBoundDevicePage(claimPage, task)) { pause('当前标签页已离开绑定的设备授权页，未领取安全凭据'); return; }
                    GM_setValue(WORKER, { runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                        heartbeat: Date.now(), paused: false, freshLogin: false });
                    const value = GM_getValue(ticketKey, null);
                    if (value?.runId === task.runId && value.sessionId === task.sessionId && value.bindingId === task.bindingId &&
                        value.boundAt === task.boundAt && value.publishedAt >= task.boundAt && secretShape(value.ticket) && value.expiresAt > Date.now()) {
                        handoff = value; break;
                    }
                    await sleep(200);
                    if (!sameBoundDevicePage(claimPage, task)) { pause('当前标签页已离开绑定的设备授权页，未领取安全凭据'); return; }
                }
                if (!handoff) { pause('未收到此标签页专属的安全凭据，已暂停'); return; }
                if (!sameBoundDevicePage(claimPage, task)) { pause('当前标签页已离开绑定的设备授权页，未领取安全凭据'); return; }
                GM_deleteValue(ticketKey);
                const claimed = await helperRequest('claim', null, { ticket: handoff.ticket });
                handoff.ticket = '';
                if (claimed.sessionId !== task.sessionId || !secretShape(claimed.capability)) throw new Error('binding');
                if (!sameBoundDevicePage(claimPage, task)) {
                    claimed.capability = ''; throw new Error('paused');
                }
                task.capability = claimed.capability;
                task.expiresAt = Date.parse(claimed.expiresAt);
                claimed.capability = '';
                await save();
                if (!sameBoundDevicePage(claimPage, task)) throw new Error('changed');
            }
            if (task.stopped || !secretShape(task.capability)) { ui.set('此标签页已转人工，请在后台检查当前会话。'); return; }
            let unknownSince = 0;
            let lastAction = null;
            let lastAt = 0;
            while (mayAct(readControl(), task, Date.now()) && !task.stopped) {
                const now = Date.now();
                const state = readControl();
                const storedWorker = GM_getValue(WORKER, null);
                const submittedWorker = task.passwordSubmission
                    ? { ...task.passwordSubmission, paused: false, freshLogin: true,
                        freshLoginAt: task.passwordSubmission.submittedAt }
                    : null;
                const confirmedWorker = workerConfirmedFreshLogin(state, submittedWorker, now) ? submittedWorker :
                    workerConfirmedFreshLogin(state, storedWorker, now) ? storedWorker : null;
                GM_setValue(WORKER, { runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                    heartbeat: now, paused: false, freshLogin: !!confirmedWorker,
                    ...(confirmedWorker ? { freshLoginAt: confirmedWorker.freshLoginAt } : {}) });
                let current = await helperRequest('current', task.capability);
                if (current.sessionId !== task.sessionId) throw new Error('binding');
                if (current.status === 'SUCCEEDED') {
                    clearPendingPassword();
                    clearSubmissionProof();
                    task.capability = ''; task.stopped = true; await save();
                    ui.set('后台已核验成功，等待队列安排下一项。'); return;
                }
                restoreSubmissionProofs(current);
                if (!mayAct(readControl(), task, Date.now()) || task.stopped) break;
                const view = snapshot();
                if (pendingPasswordInput && view.password !== pendingPasswordInput) clearPendingPassword();
                const action = decide(view, current.email, current.clientId, task);
                ui.set(`${current.email}\n只处理当前会话，安全挑战请人工完成。`);
                if (action === 'pause' || action === 'mismatch') {
                    pause(action === 'mismatch' ? 'Microsoft 显示账号不匹配，已暂停' :
                        view.consent || view.continue ? '无法证明此同意/继续页面属于当前设备码应用，已暂停' :
                            '检测到验证、错误或缺少本次设备码提交证明，已暂停'); return;
                }
                if (action === 'wait') { await sleep(2500); continue; }
                if (action === 'unknown') {
                    unknownSince ||= Date.now();
                    if (Date.now() - unknownSince > 8000) { pause('Microsoft 页面无法识别，已暂停'); return; }
                    await sleep(1000); continue;
                }
                unknownSince = 0;
                if (task.actions[action]) {
                    if (lastAction === action && Date.now() - lastAt < 8000) { await sleep(1000); continue; }
                    pause('已提交过此步骤；为避免重复提交，请人工检查'); return;
                }
                let marker = captureAction(action, view, current);
                if (!marker) { pause('无法确认提交目标、请求方法或 OAuth 应用绑定，已暂停'); return; }
                if (submittedDeviceCode(task, current.clientId) && view.identities.some((identity) => normalize(identity) === normalize(current.email))) {
                    task.identityVerified = true;
                }
                task.actions[action] = true;
                await save();
                marker = recaptureAction(marker, current, task);
                if (!marker) { pause('Microsoft 页面在操作前发生变化，已暂停'); return; }
                if (action === 'password') {
                    let credential = await helperRequest('password', task.capability);
                    try {
                        marker = recaptureAction(marker, current, task);
                        if (!marker) throw new Error('changed');
                        if (typeof credential.password !== 'string' || !credential.password) throw new Error('password');
                        pendingPasswordInput = marker.input;
                        fill(marker.input, credential.password);
                    } finally { credential.password = ''; credential = null; }
                    task.passwordSent = true;
                    task.identityVerified = true;
                    await save();
                    const finalMarker = recaptureAction(marker, current, task);
                    if (!finalMarker) {
                        fill(marker.input, '');
                        pause('密码填写后 Microsoft 页面或会话绑定发生变化，已清空密码并暂停'); return;
                    }
                    task.passwordSubmission = { runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                        clientId: current.clientId, email: current.email, submittedAt: Date.now() };
                    if (!persistSubmissionProofs(current)) {
                        task.passwordSubmission = null;
                        fill(marker.input, '');
                        pause('无法可靠保存本次密码提交证明，已清空密码并暂停'); return;
                    }
                    GM_setValue(WORKER, { runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                        heartbeat: Date.now(), paused: false, freshLogin: true,
                        freshLoginAt: task.passwordSubmission.submittedAt });
                    // The verified Microsoft button click is the last synchronous operation. Modern
                    // Microsoft pages can navigate from a click handler without dispatching a native
                    // form submit event, so waiting for submit would lose the proof during navigation.
                    try { finalMarker.target.click(); }
                    catch (error) {
                        task.passwordSubmission = null;
                        GM_setValue(WORKER, { runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                            heartbeat: Date.now(), paused: true, freshLogin: false, reason: '密码登录按钮未能执行，已暂停' });
                        throw error;
                    }
                    void save();
                } else {
                    if (action === 'device') fill(marker.input, current.userCode);
                    if (action === 'email') fill(marker.input, current.email);
                    const finalMarker = recaptureAction(marker, current, task);
                    if (!finalMarker) { pause('Microsoft 页面在提交前发生变化，已暂停'); return; }
                    if (action === 'device' || action === 'email') {
                        const proof = { runId: task.runId, sessionId: task.sessionId, bindingId: task.bindingId,
                            clientId: current.clientId, submittedAt: Date.now() };
                        if (action === 'device') task.deviceSubmission = proof;
                        else task.emailSubmission = { ...proof, email: current.email };
                        if (!persistSubmissionProofs(current)) {
                            pause(`无法可靠保存本次${action === 'device' ? '设备码' : '邮箱'}操作证明，已暂停`); return;
                        }
                    }
                    // No asynchronous boundary is permitted between proof persistence, final
                    // validation and this trusted Microsoft button click.
                    try { finalMarker.target.click(); }
                    catch (error) {
                        clearSubmissionProof();
                        throw error;
                    }
                    void save();
                }
                marker = null;
                current = null;
                lastAction = action;
                lastAt = Date.now();
                await sleep(1500);
            }
            pause('助手已暂停、后台失联或凭据到期，已转人工');
        } catch {
            pause('安全凭据或页面操作失败，已暂停；密码不会重复领取');
        }
    }
    if (location.origin === ORIGIN) {
        const bootstrap = readBootstrap(location.hash);
        if (bootstrap && history?.replaceState) history.replaceState(null, '', location.pathname);
        if (bootstrap) void bootstrapMain(bootstrap);
        else void adminMain();
    }
    else if (MS_HOSTS.includes(location.hostname)) void microsoftMain();
})();
