import { useCallback, useEffect, useRef, useState } from 'react';
import { Alert, App, Button, Card, Col, Empty, Progress, Row, Space, Statistic, Table, Tag, Typography } from 'antd';
import { CopyOutlined, DownloadOutlined, PauseCircleOutlined, PlayCircleOutlined, ReloadOutlined } from '@ant-design/icons';
import { reauthorizationApi, type ReauthorizationCandidate, type ReauthorizationSession, type ReauthorizationStatus } from '../../api';
import { requestData } from '../../utils/request';

type Outcome = 'waiting' | 'success' | 'skipped' | 'failed' | 'unsupported';
type QueueRow = ReauthorizationCandidate & { outcome: Outcome };
type Snapshot = ReauthorizationSession & { receivedAt: number };
const activeStatuses: ReauthorizationStatus[] = ['STARTING', 'PENDING', 'POLLING'];
const statusLabels: Record<ReauthorizationStatus, string> = {
    STARTING: '正在生成验证码', PENDING: '等待 Microsoft 授权', POLLING: '正在核验授权',
    SUCCEEDED: '授权成功', MISMATCH: '登录账号不匹配', DECLINED: '授权被拒绝',
    EXPIRED: '验证码已过期', CANCELLED: '已跳过', FAILED: '授权失败',
};
const outcomeLabels: Record<Outcome, string> = { waiting: '待授权', success: '成功', skipped: '跳过', failed: '失败', unsupported: '不支持' };
function until(snapshot: Snapshot, date: string, now: number) {
    return Math.max(0, Date.parse(date) - Date.parse(snapshot.serverTime) - (now - snapshot.receivedAt));
}

export default function ReauthorizationsPage() {
    const { message } = App.useApp();
    const [queue, setQueue] = useState<QueueRow[]>([]);
    const [session, setSession] = useState<Snapshot | null>(null);
    const [running, setRunning] = useState(false);
    const [busy, setBusy] = useState(false);
    const [helperActive, setHelperActive] = useState(false);
    const [loaded, setLoaded] = useState(false);
    const [now, setNow] = useState(Date.now);
    const mounted = useRef(false);
    const inFlight = useRef(false);
    const accept = useCallback((value: ReauthorizationSession) => {
        setSession({ ...value, receivedAt: Date.now() });
        if (activeStatuses.includes(value.status)) return;
        const outcome: Outcome = value.status === 'SUCCEEDED' ? 'success' : value.status === 'CANCELLED' ? 'skipped' : 'failed';
        setQueue((rows) => rows.map((row) => row.emailId === value.emailId ? { ...row, outcome, activeSession: null } : row));
        if (outcome === 'failed') setRunning(false);
    }, []);

    const runRequest = useCallback(async (operation: () => Promise<ReauthorizationSession | null>) => {
        if (inFlight.current) return;
        inFlight.current = true;
        setBusy(true);
        try {
            const value = await operation();
            if (!mounted.current) return;
            if (value) accept(value);
            else setRunning(false);
        } finally {
            inFlight.current = false;
            if (mounted.current) setBusy(false);
        }
    }, [accept]);

    const load = useCallback(async () => {
        if (inFlight.current) return;
        inFlight.current = true;
        setBusy(true);
        try {
            const candidates = await requestData<ReauthorizationCandidate[]>(reauthorizationApi.candidates, '读取待授权邮箱失败');
            if (!mounted.current || !candidates) return;
            setQueue(candidates.map((candidate) => ({ ...candidate, outcome: candidate.supported ? 'waiting' : 'unsupported' })));
            const restored = candidates.find((candidate) => candidate.activeSession)?.activeSession;
            setSession(restored ? { ...restored, receivedAt: Date.now() } : null);
            setRunning(false);
            setLoaded(true);
        } finally {
            inFlight.current = false;
            if (mounted.current) setBusy(false);
        }
    }, []);

    useEffect(() => {
        mounted.current = true;
        void load();
        const timer = window.setInterval(() => setNow(Date.now()), 1000);
        return () => { mounted.current = false; window.clearInterval(timer); };
    }, [load]);

    useEffect(() => {
        const sync = () => {
            const active = document.documentElement.getAttribute('data-gongxi-helper-busy') === 'true';
            setHelperActive(active);
            if (active) {
                setRunning(false);
            }
        };
        const observer = new MutationObserver(sync);
        observer.observe(document.documentElement, { attributes: true, attributeFilter: ['data-gongxi-helper-busy'] });
        sync();
        return () => observer.disconnect();
    }, []);

    const check = useCallback(async () => {
        if (!session || until(session, session.nextPollAt, Date.now()) > 0) return;
        await runRequest(() => requestData<ReauthorizationSession>(
            () => session.status === 'PENDING' ? reauthorizationApi.poll(session.sessionId) : reauthorizationApi.get(session.sessionId),
            '检查授权失败，请继续或重试',
        ));
    }, [session, runRequest]);

    const start = useCallback(async (emailId: number) => {
        await runRequest(() => requestData<ReauthorizationSession>(() => reauthorizationApi.start(emailId), '生成验证码失败，请刷新队列后重试'));
    }, [runRequest]);

    // One timer and one request at a time. Server-provided timestamps account
    // for clock skew; the server also enforces interval and a database claim.
    useEffect(() => {
        if (!running || busy || !loaded || helperActive) return;
        let timer: number | undefined;
        if (session && activeStatuses.includes(session.status)) {
            timer = window.setTimeout(() => { void check(); }, Math.max(250, until(session, session.nextPollAt, Date.now())));
        } else {
            const next = queue.find((row) => row.supported && row.outcome === 'waiting');
            if (next) timer = window.setTimeout(() => { void start(next.emailId); }, 250);
            else timer = window.setTimeout(() => setRunning(false), 0);
        }
        return () => { if (timer !== undefined) window.clearTimeout(timer); };
    }, [running, busy, loaded, helperActive, session, queue, check, start]);

    const skip = async () => {
        if (!session || inFlight.current) return;
        if (activeStatuses.includes(session.status)) {
            await runRequest(() => requestData<ReauthorizationSession>(() => reauthorizationApi.cancel(session.sessionId), '跳过失败，请重试'));
        } else {
            setQueue((rows) => rows.map((row) => row.emailId === session.emailId ? { ...row, outcome: 'skipped' } : row));
            setSession(null);
        }
        // Continue only after the cancellation has a confirmed terminal result.
        // A failed request pauses inside runRequest and must not auto-advance.
    };

    const retry = () => {
        if (!session) return;
        const emailId = session.emailId;
        setQueue((rows) => rows.map((row) => row.emailId === emailId ? { ...row, outcome: 'waiting' } : row));
        setRunning(true);
        void start(emailId);
    };

    const openMicrosoft = () => {
        if (!session?.userCode || !session.verificationUri) return;
        // Only this explicit click opens a tab; advancing the queue never does.
        window.open(session.verificationUriComplete || session.verificationUri, '_blank', 'noopener,noreferrer');
        if (navigator.clipboard) {
            void navigator.clipboard.writeText(session.userCode)
                .then(() => message.success('验证码已复制，请在微软页面确认当前邮箱'))
                .catch(() => message.warning('自动复制失败，请手动复制下方验证码'));
        } else void message.warning('当前浏览器无法自动复制，请手动复制验证码');
    };

    const counts = queue.reduce((result, row) => ({ ...result, [row.outcome]: result[row.outcome] + 1 }), { waiting: 0, success: 0, skipped: 0, failed: 0, unsupported: 0 });
    const active = !!session && activeStatuses.includes(session.status);
    const secondsLeft = session ? Math.ceil(until(session, session.expiresAt, now) / 1000) : 0;
    const waitSeconds = session ? Math.ceil(until(session, session.nextPollAt, now) / 1000) : 0;
    const canRetry = !!session && !active && session.status !== 'SUCCEEDED';
    const completed = counts.success + counts.skipped + counts.failed + counts.unsupported;

    return <Space orientation="vertical" size="large" style={{ width: '100%' }}>
        <div>
            <Typography.Title level={3} style={{ marginTop: 0 }}>批量重新授权</Typography.Title>
            <Typography.Paragraph type="secondary">Microsoft 仍要求每个账号确认一次。请确保微软页面登录的账号与右侧当前邮箱一致；确认后系统会核验 Graph Mail.ReadWrite 权限并进入下一个邮箱。</Typography.Paragraph>
        </div>
        <Alert type="info" showIcon title="仅显示明确要求重新登录的未禁用邮箱" description="无需粘贴令牌。仅 IMAP 分组会明确标为不支持并由自动队列跳过；请先改为可回退到 Graph 的策略，或另行完成 IMAP 授权。页面刷新后可恢复受支持账号未完成的验证码。" />
        <Card title="Tampermonkey 串行授权助手（可选）">
            <Space orientation="vertical" style={{ width: '100%' }}>
                <Button type="primary" icon={<DownloadOutlined />} href="/gongxi-mail-reauthorization.user.js" target="_blank" rel="noopener noreferrer">安装 / 更新油猴脚本</Button>
                <Typography.Text>安装 Tampermonkey 扩展并允许脚本运行后刷新此页，在右下角助手面板点击「启动」。仅支持 outlook.wujiaqiao.dpdns.org；请先用 1–2 个账号试运行。</Typography.Text>
                <Typography.Text type="secondary">助手逐项填写设备码、邮箱及保存的密码，仅在后台核验成功后等待至少 10 秒再开始下一项。MFA、验证码、密码错误、锁定、身份不符或未知页面会暂停，需人工完成；同一会话的密码仅自动领取一次。</Typography.Text>
                <Typography.Text type="secondary">暂停会将当前 Microsoft 页面转人工。若需使用下方手动操作，请点击助手「结束助手 / 切换人工」，再刷新队列。不要同时运行手动队列或多个助手；刷新后台后请关闭遗留的 Microsoft 助手标签页再重新启动。</Typography.Text>
                {helperActive && <Alert type="warning" showIcon title="助手正在管理队列，下方启动、跳过和重试已停用" description="当前进度以右下角助手面板为准。遇到验证时在 Microsoft 页面人工处理，后台仍会检查当前授权结果。" />}
            </Space>
        </Card>
        <Row gutter={[24, 24]}>
            <Col xs={24} xl={12}>
                <Card title={`待授权队列（${queue.length}）`} extra={<Button icon={<ReloadOutlined />} disabled={busy || running} onClick={() => void load()}>刷新队列</Button>}>
                    <Row gutter={12}>
                        <Col flex="1"><Statistic title="待授权" value={counts.waiting} /></Col>
                        <Col flex="1"><Statistic title="成功" value={counts.success} /></Col>
                        <Col flex="1"><Statistic title="跳过" value={counts.skipped} /></Col>
                        <Col flex="1"><Statistic title="失败" value={counts.failed} /></Col>
                        <Col flex="1"><Statistic title="不支持" value={counts.unsupported} /></Col>
                    </Row>
                    <Progress percent={queue.length ? Math.round(completed / queue.length * 100) : 0} style={{ margin: '16px 0' }} />
                    <Table<QueueRow> size="small" rowKey="emailId" dataSource={queue} loading={!loaded && busy}
                        pagination={{ pageSize: 10, showSizeChanger: false }} scroll={{ x: 420 }}
                        columns={[
                            { title: '邮箱', dataIndex: 'email', render: (value: string, row) => <Space orientation="vertical" size={0}><Typography.Text strong={row.emailId === session?.emailId}>{value}</Typography.Text><Typography.Text type="secondary">{row.groupName || '未分组'}</Typography.Text>{row.unsupportedReason && <Typography.Text type="danger">{row.unsupportedReason}</Typography.Text>}</Space> },
                            { title: '状态', width: 95, render: (_value: unknown, row) => <Tag color={row.outcome === 'success' ? 'success' : row.outcome === 'failed' || row.outcome === 'unsupported' ? 'error' : row.emailId === session?.emailId ? 'processing' : 'default'}>{row.emailId === session?.emailId && active ? '当前账号' : outcomeLabels[row.outcome]}</Tag> },
                        ]} />
                </Card>
            </Col>
            <Col xs={24} xl={12}>
                <Card title="当前授权" extra={session && <Tag color={session.status === 'SUCCEEDED' ? 'success' : active ? 'processing' : 'warning'}>{statusLabels[session.status]}</Tag>}>
                    {session ? <Space orientation="vertical" size="large" style={{ width: '100%' }}>
                        <Typography.Title level={4} copyable={{ text: session.email, tooltips: ['复制邮箱', '邮箱已复制'] }} style={{ margin: 0, overflowWrap: 'anywhere' }}>{session.email}</Typography.Title>
                        {session.userCode && <div style={{ textAlign: 'center', padding: 24, background: '#f0f5ff', borderRadius: 8 }}>
                            <Typography.Paragraph type="secondary">Microsoft 设备验证码</Typography.Paragraph>
                            <Typography.Text copyable={{ text: session.userCode }} style={{ fontSize: 32, fontFamily: 'monospace', fontWeight: 700, letterSpacing: 3 }}>{session.userCode}</Typography.Text>
                            <div style={{ marginTop: 16 }}>有效期剩余 {Math.floor(secondsLeft / 60)}:{String(secondsLeft % 60).padStart(2, '0')}</div>
                        </div>}
                        {active && <Button type="primary" size="large" block icon={<CopyOutlined />} disabled={helperActive || !session.userCode || secondsLeft <= 0} onClick={openMicrosoft}>复制验证码并打开微软授权页面</Button>}
                        {session.errorMessage && <Alert type={active ? 'warning' : 'error'} showIcon title={session.errorMessage} description={session.authorizedEmail ? `Microsoft 返回账号：${session.authorizedEmail}` : undefined} />}
                        {session.status === 'SUCCEEDED' && <Alert type="success" showIcon title="账号与 Graph Mail.ReadWrite 权限已核验并保存，邮箱已恢复为正常状态" />}
                        <Space wrap>
                            <Button icon={running ? <PauseCircleOutlined /> : <PlayCircleOutlined />} disabled={helperActive || !running && !active && counts.waiting === 0} onClick={() => setRunning((value) => !value)}>{running ? '暂停' : '继续'}</Button>
                            <Button loading={busy} disabled={!active || busy || waitSeconds > 0} onClick={() => void check()}>{waitSeconds > 0 ? `${waitSeconds} 秒后可检查` : '立即检查'}</Button>
                            <Button disabled={helperActive || busy || session.status === 'SUCCEEDED' || session.status === 'CANCELLED'} onClick={() => void skip()}>跳过</Button>
                            <Button disabled={helperActive || busy || !canRetry} onClick={retry}>重试当前账号</Button>
                        </Space>
                        {!running && active && <Typography.Text type="secondary">已暂停。微软确认完成后，请点击继续或立即检查。</Typography.Text>}
                    </Space> : <Space orientation="vertical" size="large" style={{ width: '100%', textAlign: 'center' }}>
                        <Empty description={loaded && !queue.length ? '当前没有需要重新授权的邮箱' : '准备开始逐个授权'} />
                        <Button type="primary" size="large" loading={busy} disabled={helperActive || !loaded || !counts.waiting} onClick={() => setRunning(true)}>开始批量授权</Button>
                    </Space>}
                </Card>
            </Col>
        </Row>
    </Space>;
}
