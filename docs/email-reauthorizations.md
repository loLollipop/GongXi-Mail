# Microsoft 邮箱批量重新授权

超级管理员在「邮箱管理」后的「批量重新授权」页面启动队列。默认手动模式需逐个在 Microsoft 页面确认；也可安装页面提供的 Tampermonkey 助手处理明确的标准登录步骤。device code 加密存储，access token 仅在服务端请求内存中使用，refresh token 核验成功后加密保存。

## 可选 Tampermonkey 助手

安装 Tampermonkey 扩展，在后台 `/reauthorizations` 点击「安装 / 更新油猴脚本」，确认脚本权限后刷新页面。需要最新版 Chrome / Edge、扩展允许用户脚本运行，以及生产域 `https://outlook.wujiaqiao.dpdns.org`。脚本仅连接此固定后台域；不适用于 localhost 或其他域名。点击右下角助手面板「启动 / 继续队列」，保持后台标签页打开。首次必须先用 1–2 个已获授权的测试邮箱试运行真实 Microsoft 页面。

助手从后台候选队列取当前项，打开一个 Microsoft 设备授权标签页，串行填写设备码、邮箱及已保存密码，并在身份匹配后处理「保持登录」的否定选项。设备码只有在该绑定标签页实际观察到表单 `submit` 事件后才留下后续登录步骤所需的证明；密码页还必须显示与当前会话一致的邮箱，密码才会被领取一次。只有服务端确认 Graph `Mail.ReadWrite` 与 `/me` 身份匹配成功后才计为成功，再等待至少 10 秒进入下一项。Microsoft 页面出现成功文字不能直接推进队列。

助手不会自动点击 OAuth 「同意 / 允许」或「继续」按钮；这些控件与登录输入混合出现时也始终优先暂停。helper `current` 仅向已绑定标签页提供当前会话预期的非密密 OAuth `client_id`；设备码提交证明会绑定该值。后续填写邮箱、密码、选择其他账号或否定「保持登录」之前，当前 URL 和按浏览器规则计算的实际提交目标（包括按钮 `formaction` / `formmethod` 覆盖）都必须是受支持的 HTTPS Microsoft 域名、显式携带 `client_id`，且所有值与服务端预期值严格一致；密码表单还必须使用 POST。缺失、不匹配或操作期间发生变化一律清空已填密码并暂停转人工，避免同一标签页跨导航后误向其他应用提交凭据。

MFA、验证码、错误密码、账号锁定/恢复、未知页面、显示账号不匹配或网络异常都会暂停。助手不会解验证码、自动处理 MFA、伪造设备、绕过风控或重试密码。Microsoft 页面变化可能导致暂停；在该页面人工完成后，后台仍会核验当前授权，点击「继续队列」才恢复后续项。当前项一旦转人工，不再自动填写。手动操作后台前应点击「结束助手 / 切换人工」，然后刷新队列；失败的会话按原有重试流程处理。刷新或关闭后台会使 Microsoft 助手停止，请关闭遗留助手标签页再启动新一轮。

安全边界：管理员 JWT 只从后台 origin 的 localStorage 读取并发送同源管理请求，不写 GM 存储、不进入 Microsoft origin。每个会话仅签发一次最多 90 秒的随机 ticket；Redis 只存 SHA-256 摘要及 sessionId、emailId、adminId、tokenVersion、UA 摘要、有效期。Microsoft 侧一次性领取 ticket 后，获得最长 5 分钟且不超过 device session 有效期的单会话 capability。capability 仅可读取当前会话（包括用于页面流程绑定的非密密 `client_id`）和领取一次密码，不能列举、指定账号、访问管理 API、轮询 Microsoft token 或进入下一项。

密码只有在脚本识别并核对当前账号的密码输入页后才通过 capability 原子领取一次；领取前、填入前和最终点击前都会重新采集页面，复核运行/会话/标签页绑定、目标邮箱、`client_id` 以及同一输入框、按钮和表单。最终复核与点击之间没有异步等待。密码只在内存中填入目标密码输入框，随后清空脚本引用；不写 GM、localStorage、剪贴板、URL、状态面板或日志。GM 中的一次性 ticket 领取前即删除，capability 存在 Tampermonkey 的当前标签页隔离状态中；暂停/失联时清除，服务端 TTL 限制始终生效。密码响应丢失或解析失败后不再次发放，需人工处理或取消并创建新会话。

助手要求可用的 Redis **6.2+**（`GETDEL` 与 Lua 脚本），没有内存回退；请配置 `REDIS_URL`，保持同一部署的所有服务实例使用同一个 Redis，避免运行期间清空助手键或让键被提前驱逐。Redis 不可用时关闭助手发放功能，手动授权仍可使用。反向代理需保留生产 Host，并禁止记录 helper 请求/响应 body、自定义 capability 请求头和管理 Authorization；API 均返回 `Cache-Control: no-store`。服务端复核账号版本、会话状态、管理员状态与 UA，ticket 消费/密码领取为原子操作；同一管理员的助手会话有 Redis 互斥与成功后 10 秒间隔。审计仅记录事件类型与管理员/邮箱/会话 ID。

本助手不增加数据库字段或迁移；新增 `/api/reauthorization-helper/{claim,current,password}` 和受超级管理员 JWT 保护的 `POST /admin/email-reauthorizations/:id/helper-ticket`。部署时一并更新 server、web 和 userscript；旧脚本应在安装入口手动更新。代理禁止重定向 helper 请求。浏览器扩展及 Microsoft 页面能接触输入框中的密码，因此应使用可信浏览器环境。后端测试使用内存 Redis 语义替身与模拟数据库，未代替真实 Redis 原子行为、多实例和 Microsoft 页面试运行。

## 部署与数据库

此仓库的 Docker 启动命令使用 `prisma db push --skip-generate`，没有迁移历史。本次保持这个部署约定：在更新服务之前备份数据库，生成新 Prisma Client，并在维护窗口执行 `prisma db push`。不得使用 `--accept-data-loss`。本次 schema 变更仅增加 `email_accounts.token_version`（默认 0）、授权会话表、枚举、外键和索引，不删除现有数据。不要直接在已有、未建立迁移基线的数据库执行 `prisma migrate deploy`。

现有邮箱从版本 0 开始；所有新版本的手工修改、导入和 token 轮换都会增加版本。部署时停止旧实例及其定时刷新任务，完成 schema 更新后只运行新实例；旧二进制不遵守版本 CAS，会破坏并发保护。回退代码时先停止授权和 token 刷新任务，确认没有在途请求；新增列/表可保留。

数据库唯一的可空 `active_email_id` 保证同一邮箱只有一个活动会话，不依赖单进程锁。轮询通过数据库状态与 claim 比较更新串行化；token 写入和会话成功状态位于同一个事务，并对邮箱版本、身份和禁用状态进行校验。请求超时 15 秒，轮询占用超过 60 秒后会话失败，需重新开始，避免重复兑换已消费的设备码。

到期会话在访问授权接口时清理，终态立即清除设备码与用户验证码；终态记录保留 30 天并按需删除。无后台轮询 worker。页面刷新可恢复仍活动的会话，点击继续后恢复轮询。本轮成功/跳过/失败计数在刷新队列时重置。

## Microsoft 应用配置

使用各邮箱当前 `clientId`，应用必须支持个人 Microsoft 账号与 public client/device-code flow。默认请求 `openid profile offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Mail.ReadWrite`，仅包含 Graph 资源且不同时请求 `Mail.Read`。token 响应必须明确包含大小写不敏感的 `Mail.ReadWrite` 或完整 Graph URI scope，否则以 `GRAPH_MAIL_READWRITE_SCOPE_MISSING` 失败且不保存 refresh token。`GET /me` 的 `mail` 或 `userPrincipalName` 必须与目标邮箱大小写不敏感相等；无法证明别名归属时拒绝保存。客户端配置或租户策略不允许 device flow 时页面显示失败，不能以跳过身份核验的方式绕过。

本功能恢复 Graph 读写授权，因为现有 `processMailbox` 会删除邮件，必须具备 `Mail.ReadWrite`。本功能不会申请 IMAP 权限；`IMAP_ONLY` 分组会在候选列表明确标为不支持，服务端也会以 `REAUTHORIZATION_STRATEGY_UNSUPPORTED` 拒绝启动。此类账号需先改为可回退到 Graph 的策略，或另行完成 IMAP 授权。自动 token 刷新会跳过明确要求用户重新登录的账号，单账号手工刷新仍可重试。

## 上线核验

本地自动测试使用模拟 Microsoft 响应和数据库操作，不连接真实账号。上线前在独立测试库检验 `prisma db push`、同邮箱并发启动唯一性、跨实例轮询 CAS、取消与兑换竞争，以及 Microsoft 应用的 public client 配置；再使用获授权的测试邮箱验证真实登录、错误账号拒绝、成功后的 Graph 读信和页面刷新恢复。不要把令牌或 device code 放进测试日志。
