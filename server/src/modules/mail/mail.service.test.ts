import assert from 'node:assert/strict';
import test from 'node:test';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
process.env.JWT_SECRET = 'test-jwt-secret-for-mail-fallback-0000';

const { AppError } = await import('../../plugins/error.js');
const { default: Imap } = await import('node-imap');
const {
    applyXoauth2LoginCompatibility,
    mailService,
    normalizeImapConnectionError,
    ProtocolConsentRequiredError,
} = await import('./mail.service.js');

const credentials = {
    id: 7,
    email: 'test@outlook.com',
    clientId: 'client-id',
    refreshToken: 'refresh-token',
    tokenVersion: 4,
    autoAssigned: false,
};

void test('GRAPH_FIRST falls back to IMAP for AADSTS65001-style protocol consent failure', async (t) => {
    const graph = t.mock.method(mailService, 'getGraphAccessToken', async () => {
        throw new ProtocolConsentRequiredError('invalid_grant: AADSTS65001: REAUTHORIZATION_REQUIRED');
    });
    const imapToken = t.mock.method(mailService, 'getImapAccessToken', async () => 'imap-access-token');
    const imap = t.mock.method(mailService, 'getEmailsViaImap', async () => [{
        id: 'message-1', from: 'from@example.com', subject: 'subject', text: 'body', html: '', date: '',
    }]);

    const result = await mailService.getEmails(
        { ...credentials, fetchStrategy: 'GRAPH_FIRST' },
        { mailbox: 'INBOX' },
    );

    assert.equal(result.method, 'imap');
    assert.equal(result.count, 1);
    assert.equal(graph.mock.callCount(), 1);
    assert.equal(imapToken.mock.callCount(), 1);
    assert.equal(imap.mock.callCount(), 1);
});

void test('IMAP_FIRST symmetrically falls back to Graph for protocol consent failure', async (t) => {
    const imapToken = t.mock.method(mailService, 'getImapAccessToken', async () => {
        throw new ProtocolConsentRequiredError('invalid_grant: AADSTS65001: REAUTHORIZATION_REQUIRED');
    });
    const graphToken = t.mock.method(mailService, 'getGraphAccessToken', async () => ({
        accessToken: 'graph-access-token',
        hasMailRead: true,
    }));
    const graph = t.mock.method(mailService, 'getEmailsViaGraphApi', async () => []);

    const result = await mailService.getEmails(
        { ...credentials, fetchStrategy: 'IMAP_FIRST' },
        { mailbox: 'INBOX' },
    );

    assert.equal(result.method, 'graph_api');
    assert.equal(imapToken.mock.callCount(), 1);
    assert.equal(graphToken.mock.callCount(), 1);
    assert.equal(graph.mock.callCount(), 1);
});

void test('explicit credential invalidation remains terminal and skips fallback', async (t) => {
    const graph = t.mock.method(mailService, 'getGraphAccessToken', async () => {
        throw new AppError('REAUTHORIZATION_REQUIRED', 'invalid_grant: AADSTS70000: REAUTHORIZATION_REQUIRED', 409);
    });
    const imap = t.mock.method(mailService, 'getImapAccessToken', async () => 'imap-access-token');

    await assert.rejects(
        mailService.getEmails(
            { ...credentials, fetchStrategy: 'GRAPH_FIRST' },
            { mailbox: 'INBOX' },
        ),
        (error: unknown) => error instanceof AppError && error.code === 'REAUTHORIZATION_REQUIRED',
    );
    assert.equal(graph.mock.callCount(), 1);
    assert.equal(imap.mock.callCount(), 0);
});

function createCapabilityReader(capabilities: string[]) {
    return {
        serverSupports(capability: string): boolean {
            return capabilities.includes(capability);
        },
    };
}

void test('XOAUTH2 compatibility masks LOGINDISABLED when the server advertises XOAUTH2', () => {
    const imap = createCapabilityReader(['IMAP4REV1', 'LOGINDISABLED', 'AUTH=XOAUTH2']);

    applyXoauth2LoginCompatibility(imap, 'configured-xoauth2');

    assert.equal(imap.serverSupports('LOGINDISABLED'), false);
    assert.equal(imap.serverSupports('AUTH=XOAUTH2'), true);
    assert.equal(imap.serverSupports('IMAP4REV1'), true);
});

void test('XOAUTH2 compatibility reaches the real node-imap authentication command', () => {
    type LoginHarness = InstanceType<typeof Imap> & {
        _caps: string[];
        _login(): void;
        _enqueue(command: string, callback: (error?: Error) => void): void;
    };
    const authString = 'configured-xoauth2';
    const imap = new (Imap as typeof Imap)({
        user: 'test@outlook.com',
        password: '',
        xoauth2: authString,
        host: 'outlook.office365.com',
        port: 993,
        tls: true,
    }) as LoginHarness;
    const commands: string[] = [];
    imap.state = 'connected';
    imap._caps = ['IMAP4REV1', 'LOGINDISABLED', 'AUTH=XOAUTH2'];
    imap._enqueue = (command, callback) => {
        commands.push(command);
        if (command === 'CAPABILITY') callback();
    };

    applyXoauth2LoginCompatibility(imap, authString);
    imap._login();

    assert.deepEqual(commands, [
        'CAPABILITY',
        `AUTHENTICATE XOAUTH2 ${authString}`,
    ]);
});

void test('XOAUTH2 compatibility preserves LOGINDISABLED without server XOAUTH2 capability', () => {
    const imap = createCapabilityReader(['IMAP4REV1', 'LOGINDISABLED']);

    applyXoauth2LoginCompatibility(imap, 'configured-xoauth2');

    assert.equal(imap.serverSupports('LOGINDISABLED'), true);
    assert.equal(imap.serverSupports('AUTH=XOAUTH2'), false);
});

void test('XOAUTH2 compatibility leaves capabilities unchanged without XOAUTH2 configuration', () => {
    const imap = createCapabilityReader(['IMAP4REV1', 'LOGINDISABLED', 'AUTH=XOAUTH2']);

    applyXoauth2LoginCompatibility(imap, undefined);

    assert.equal(imap.serverSupports('LOGINDISABLED'), true);
    assert.equal(imap.serverSupports('AUTH=XOAUTH2'), true);
    assert.equal(imap.serverSupports('IMAP4REV1'), true);
});

void test('IMAP authentication errors become a safe actionable AppError', () => {
    const nativeError = Object.assign(
        new Error('AUTHENTICATE failed for secret-token'),
        { source: 'authentication' },
    );

    const error = normalizeImapConnectionError(nativeError);

    assert.ok(error instanceof AppError);
    assert.equal(error.code, 'IMAP_AUTHENTICATION_FAILED');
    assert.equal(error.statusCode, 409);
    assert.equal(
        error.message,
        'IMAP authentication failed. Reauthorize the email account and confirm IMAP access is enabled.',
    );
    assert.doesNotMatch(error.message, /secret-token/);
});

void test('non-authentication IMAP errors retain their original identity', () => {
    const nativeError = Object.assign(new Error('socket closed'), { source: 'socket' });

    assert.equal(normalizeImapConnectionError(nativeError), nativeError);
});
