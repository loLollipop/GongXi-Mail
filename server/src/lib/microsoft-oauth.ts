/** Keep upstream response bodies out of logs, database errors and public APIs. */
export function requiresReauthorization(message: string | null | undefined): boolean {
    if (!message) return false;
    return /\bAADSTS(?:70000|700082|700084|50173|50076|50079|65001)\b|reauthorization_required|unauthorized[\s_-]*or[\s_-]*expired|user\s+must\s+sign\s+in|tokens?.{0,40}revoked|revoked.{0,40}tokens?|refresh[\s_-]+token.{0,40}expired|expired.{0,40}refresh[\s_-]+token|interaction_required|login_required|consent_required/i.test(message);
}

const knownErrors = new Set([
    'invalid_grant', 'invalid_client', 'invalid_scope', 'unauthorized_client',
    'authorization_pending', 'slow_down', 'authorization_declined', 'access_denied',
    'expired_token', 'bad_verification_code', 'interaction_required', 'login_required',
    'consent_required', 'temporarily_unavailable', 'server_error', 'invalid_request',
]);

export type MicrosoftOAuthFailureKind =
    | 'PROTOCOL_CONSENT_REQUIRED'
    | 'REAUTHORIZATION_REQUIRED'
    | 'REQUEST_FAILED';

export interface MicrosoftOAuthError {
    code: string;
    message: string;
    reauthorizationRequired: boolean;
    kind: MicrosoftOAuthFailureKind;
}

export function microsoftError(body: unknown, httpStatus?: number): MicrosoftOAuthError {
    const record = body && typeof body === 'object' ? body as Record<string, unknown> : {};
    const rawCode = typeof record.error === 'string' ? record.error.toLowerCase() : '';
    const description = typeof record.error_description === 'string' ? record.error_description : '';
    const numericCode: unknown = Array.isArray(record.error_codes) ? record.error_codes[0] : undefined;
    const aadsts = description.match(/\bAADSTS\d{5,9}\b/i)?.[0].toUpperCase()
        ?? (typeof numericCode === 'number' && Number.isInteger(numericCode) && numericCode >= 10000 && numericCode <= 999999999 ? `AADSTS${numericCode}` : undefined);
    const code = knownErrors.has(rawCode) ? rawCode : 'microsoft_oauth_error';
    const reauthorizationRequired = requiresReauthorization(`${rawCode} ${description} ${aadsts ?? ''}`);
    const kind: MicrosoftOAuthFailureKind = aadsts === 'AADSTS65001'
        ? 'PROTOCOL_CONSENT_REQUIRED'
        : reauthorizationRequired ? 'REAUTHORIZATION_REQUIRED' : 'REQUEST_FAILED';
    // Only allowlisted codes and a fixed explanation; never echo arbitrary text.
    return {
        code,
        reauthorizationRequired,
        kind,
        message: [code, aadsts, httpStatus ? `HTTP ${httpStatus}` : '',
            reauthorizationRequired ? 'REAUTHORIZATION_REQUIRED: User must sign in again' : 'Microsoft request failed']
            .filter(Boolean).join(': '),
    };
}

export async function readMicrosoftError(response: Response) {
    const body: unknown = await response.json().catch(() => null);
    return microsoftError(body, response.status);
}
