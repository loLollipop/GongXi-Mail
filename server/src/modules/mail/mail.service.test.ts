import assert from 'node:assert/strict';
import test from 'node:test';

process.env.DATABASE_URL = 'postgresql://test:test@127.0.0.1:1/test';
process.env.ENCRYPTION_KEY = '01234567890123456789012345678901';
process.env.JWT_SECRET = 'test-jwt-secret-for-mail-fallback-0000';

const { AppError } = await import('../../plugins/error.js');
const { mailService, ProtocolConsentRequiredError } = await import('./mail.service.js');

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
