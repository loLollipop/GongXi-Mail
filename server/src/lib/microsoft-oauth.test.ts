import assert from 'node:assert/strict';
import test from 'node:test';
import { microsoftError, requiresReauthorization } from './microsoft-oauth.js';

void test('classifies explicit Microsoft reauthorization errors without treating network or access-token expiry as revocation', () => {
    for (const value of ['AADSTS70000: invalid grant', 'Unauthorized-or-expired', 'UnauthorizedOrExpired',
        'The user must sign in again', 'Refresh token has expired', 'Tokens have been revoked',
        'AADSTS50173', 'AADSTS65001: The user or administrator has not consented',
        'interaction_required', 'REAUTHORIZATION_REQUIRED']) {
        assert.equal(requiresReauthorization(value), true, value);
    }
    for (const value of [null, '', 'HTTP 503', 'invalid_client', 'invalid_scope', 'invalid_grant',
        'access token has expired', 'No refresh_token in response', 'ETIMEDOUT', 'Failed to decrypt refresh token']) {
        assert.equal(requiresReauthorization(value), false, String(value));
    }
});

void test('upstream descriptions and unknown codes never echo secrets or newlines', () => {
    const error = microsoftError({ error: 'invalid_grant', error_description: 'AADSTS70000\r\nrefresh_token=secret-token device_code=secret-code' }, 400);
    assert.equal(error.reauthorizationRequired, true);
    assert.match(error.message, /AADSTS70000/);
    assert.doesNotMatch(error.message, /secret|\r|\n/);
    assert.equal(microsoftError({ error: 'secret-token' }).code, 'microsoft_oauth_error');
    assert.doesNotMatch(microsoftError({ error: 'secret-token' }).message, /secret/);
    assert.equal(microsoftError({ error: 'invalid_grant', error_codes: [70000] }).reauthorizationRequired, true);
    assert.equal(microsoftError({ error: 'invalid_grant', error_codes: [70000] }).kind, 'REAUTHORIZATION_REQUIRED');
    assert.equal(microsoftError({ error: 'invalid_grant', error_codes: [65001] }).reauthorizationRequired, true);
    assert.equal(microsoftError({ error: 'invalid_grant', error_codes: [65001] }).kind, 'PROTOCOL_CONSENT_REQUIRED');
    assert.equal(microsoftError({ error: 'invalid_grant', error_description: 'AADSTS65001: consent required' }).reauthorizationRequired, true);
    assert.equal(microsoftError({ error: 'invalid_grant', error_description: 'AADSTS65001: consent required' }).kind, 'PROTOCOL_CONSENT_REQUIRED');
    assert.equal(microsoftError({ error: 'consent_required' }).kind, 'REAUTHORIZATION_REQUIRED');
    assert.equal(microsoftError({ error: 'temporarily_unavailable' }).kind, 'REQUEST_FAILED');
});
