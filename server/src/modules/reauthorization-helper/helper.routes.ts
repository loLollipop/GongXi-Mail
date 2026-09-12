import type { FastifyPluginAsync } from 'fastify';
import { z } from 'zod';
import { AppError } from '../../plugins/error.js';
import { HELPER_ORIGIN, helperSecret, helperService } from './helper.service.js';

const empty = z.object({}).strict();
const helperRoutes: FastifyPluginAsync = async (fastify) => {
    fastify.addHook('onRequest', async (request, reply) => {
        reply.header('Cache-Control', 'no-store').header('Pragma', 'no-cache');
        // GM requests are anonymous; a management JWT is never an alternative.
        if (request.hostname !== new URL(HELPER_ORIGIN).hostname || request.headers.authorization ||
            request.headers.cookie || Object.keys(request.query as object).length) {
            throw new AppError('HELPER_INVALID', '助手请求无效', 401);
        }
    });
    fastify.setErrorHandler((error, _request, reply) => {
        const status = error instanceof AppError ? error.statusCode : error instanceof z.ZodError ? 400 : 503;
        return reply.status(status).send({ success: false, error: {
            code: error instanceof AppError ? error.code : 'HELPER_ERROR',
            message: status === 400 ? '助手请求参数无效' : '助手已停止，请回后台检查当前会话或人工继续',
        } });
    });
    fastify.post('/claim', { bodyLimit: 1024 }, async (request) => {
        const { ticket } = z.object({ ticket: helperSecret }).strict().parse(request.body);
        return { success: true, data: await helperService.claim(ticket, request.headers['user-agent'] ?? '') };
    });
    fastify.post('/current', { bodyLimit: 1024 }, async (request) => {
        empty.parse(request.body);
        const secret = helperSecret.parse(request.headers['x-reauthorization-capability']);
        return { success: true, data: await helperService.current(secret, request.headers['user-agent'] ?? '') };
    });
    fastify.post('/password', { bodyLimit: 1024 }, async (request) => {
        empty.parse(request.body);
        const secret = helperSecret.parse(request.headers['x-reauthorization-capability']);
        return { success: true, data: await helperService.password(secret, request.headers['user-agent'] ?? '',
            (event) => request.log.info(event, '授权助手凭据审计')) };
    });
};
export default helperRoutes;
