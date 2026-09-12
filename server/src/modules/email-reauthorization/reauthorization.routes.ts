import type { FastifyPluginAsync } from 'fastify';
import { z } from 'zod';
import { reauthorizationService } from './reauthorization.service.js';
import { AppError } from '../../plugins/error.js';
import { helperService } from '../reauthorization-helper/helper.service.js';

const sessionParams = z.object({ id: z.string().uuid() });

const reauthorizationRoutes: FastifyPluginAsync = async (fastify) => {
    fastify.addHook('preHandler', fastify.authenticateJwt);
    fastify.addHook('preHandler', fastify.requireSuperAdmin);
    fastify.addHook('onRequest', async (_request, reply) => { reply.header('Cache-Control', 'no-store'); });
    // Prisma errors can include query arguments. Never pass a secret-bearing
    // operation's raw exception into the application's generic error logger.
    fastify.setErrorHandler((error, _request, reply) => {
        const status = error instanceof AppError ? error.statusCode : error instanceof z.ZodError ? 400 : 503;
        return reply.status(status).send({ success: false, error: {
            code: error instanceof AppError ? error.code : 'REAUTHORIZATION_ERROR',
            message: error instanceof AppError ? error.message : status === 400 ? '请求参数无效' : '授权服务暂时不可用，请重试',
        } });
    });
    fastify.get('/candidates', async () => ({ success: true, data: await reauthorizationService.candidates() }));
    fastify.post('/start', async (request) => {
        const { emailId } = z.object({ emailId: z.number().int().positive() }).strict().parse(request.body);
        return { success: true, data: await reauthorizationService.start(emailId, request.user!.id) };
    });
    fastify.post('/:id/helper-ticket', async (request) => {
        z.object({}).strict().parse(request.body);
        return { success: true, data: await helperService.issue(sessionParams.parse(request.params).id,
            request.user!.id, request.headers['user-agent'] ?? '',
            (event) => request.log.info(event, '授权助手凭据审计')) };
    });
    fastify.get('/:id', async (request) => ({ success: true,
        data: await reauthorizationService.get(sessionParams.parse(request.params).id) }));
    fastify.post('/:id/poll', async (request) => ({ success: true,
        data: await reauthorizationService.poll(sessionParams.parse(request.params).id) }));
    fastify.post('/:id/cancel', async (request) => ({ success: true,
        data: await reauthorizationService.cancel(sessionParams.parse(request.params).id) }));
};

export default reauthorizationRoutes;
