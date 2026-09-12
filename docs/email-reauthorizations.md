# Microsoft 邮箱批量重新授权

超级管理员在「邮箱管理」后的「批量重新授权」页面启动队列。默认手动模式需逐个在 Microsoft 页面确认；也可安装页面提供的 Tampermonkey 助手处理明确的标准登录步骤。device code 加密存储，access token 仅在服务端请求内存中使用，refresh token 核验成功后加密保存。

## 可选 Tampermonkey 助手

兼容旧版个人账户同意页时，“明确应用名称组件”也包括同意表单内唯一且文字精确的 Thunderbird 应用链接，但页面标题必须同时是受支持的通用授权标题；正文或表单外的同名链接不能作为证明。

安装 Tampermonkey 扩展，在后台 `/reauthorizations` 点击「安装 / 更新油猴脚本」，确认脚本权限后刷新页面。需要最新版 Chrome / Edge、扩展允许用户脚本运行，以及生产域 `https://outlook.wujiaqiao.dpdns.org`。脚本仅连接此固定后台域；不适用于 localhost 或其他域名。为避免 Microsoft 旧会话介入，建议在专用无痕窗口打开后台，并允许 Tampermonkey 在无痕模式运行；脚本无法自行创建或强制无痕窗口。点击右下角助手面板「启动 / 继续队列」，保持后台标签页打开。首次必须先用 1–2 个已获授权的测试邮箱试运行真实 Microsoft 页面。

助手从后台候选队列取当前项，打开该会话由 Microsoft 实时返回的设备授权地址（当前可能为 `https://www.microsoft.com/link`，不应硬编码为旧 `/devicelogin`），先填写设备码，再依次选择明确的「通过其他 Microsoft 帐户登录 / 使用其他帐户登录」、填写队列邮箱及已保存密码，只在验证邮箱页点击明确的「使用密码」，并在身份匹配后处理新版或旧版「保持登录」页面的明确否定选项。设备码、邮箱与密码只有在绑定标签页完成最终页面、会话、表单和提交目标复核并亲自调用对应 Microsoft 登录按钮时才留下后续步骤所需的证明；密码页还必须显示与当前会话一致的邮箱，密码才会被领取一次。只有服务端确认 Graph `Mail.ReadWrite` 与 `/me` 身份匹配成功，且 Microsoft 标签页已证明本轮亲自执行了队列邮箱和密码登录按钮后才计为助手成功，再等待至少 10 秒进入下一项。Microsoft 页面出现成功文字不能直接推进队列。

助手只会在完整设备码、邮箱、密码提交证明严格绑定同一 run、session、binding、`client_id` 与队列邮箱且时间有序后，自动点击 Thunderbird OAuth 页唯一、可见、可用且文字精确为「Accept / 接受」的按钮；「Reject / 拒绝」永不成为操作目标，通用「Allow / Yes / 继续」也不会自动点击。Thunderbird 官方 OAuth `client_id` 与服务端固定的 scopes 是主要授权边界；页面还必须只有一个受支持的可见应用名称组件，且其非空文字精确显示 Thunderbird，标题、正文或任意链接中的 Thunderbird 字样都不能充当应用名证明。权限必须位于唯一受支持的完整权限容器中，容器文本须由全部可见权限条目完整覆盖，条目只能精确匹配服务端固定集合对应的邮件读写权限及已知无害说明；同时，整个同意页除该容器、应用名、当前身份和精确白名单内的 Microsoft 权限标题、服务条款/隐私声明说明及标准控件文字外，不得存在未覆盖文案。应用名组件为空、重复或相互矛盾，权限容器缺失或重复、权限只提取到一部分、容器外出现额外权限文案、固定集合以外的 scope、未知权限条目、`Mail.Send`、联系人、日历、文件、管理员等额外权限，以及任何无法确认完整性的布局都会暂停。登录输入与同意/继续控件混合出现时始终优先暂停。helper `current` 仅向已绑定标签页提供当前会话预期的非敏感 OAuth `client_id`；设备码提交证明会绑定该值。Microsoft 在设备码提交后通常用服务端流程状态传递应用上下文，后续登录 URL 和表单 action 不保证重复携带 `client_id`。因此助手不要求 consent 页重复该参数，但 URL、表单目标或可见/隐藏字段一旦显式出现 `client_id`，必须全部与服务端预期值一致；一旦出现 `scope`，只能属于固定的 `openid profile offline_access User.Read Mail.ReadWrite` 集合并且必须包含 `Mail.ReadWrite`。所有操作仍限定在由一次性票据绑定的原 Microsoft 标签页和受支持的 HTTPS Microsoft 域名，邮箱、密码及 OAuth 接受表单必须使用 POST。接受按钮、页面 URL、表单对象、action 或 method（包括按钮动态覆盖值）在点击前发生变化都会暂停。

设备码后如果出现已缓存账号选择器，助手只点击明确的「使用其他账号」，不选中任何已登录账号。Microsoft 如果先展示无密码方式，助手只在唯一受支持且标题精确为“验证你的电子邮件 / Verify your email”的证明选择页点击明确的「使用密码 / Use your password」替代入口；恢复账号、captcha、Authenticator、MFA、账号锁定、异常登录、安全代码或其他硬挑战即使同时显示“使用密码”也始终优先暂停。助手不会填写恢复邮箱、发送或读取验证码，也不会绕过 MFA。密码只会在同一绑定流程中由助手先提交了当前队列邮箱，然后密码页的账号横幅（包括新版 `#bannerText`）明确显示该邮箱时领取。如果 Microsoft 直接进入某个缓存账号的密码页、自动使用了旧会话，或未提供「使用其他账号」/全新邮箱输入步骤，助手不会领取密码，而是暂停转人工。

MFA、验证码、错误密码、账号锁定/恢复、未知页面、显示账号不匹配或网络异常都会暂停。助手不会解验证码、自动处理 MFA、伪造设备、绕过风控或重试密码。Microsoft 页面变化可能导致暂停；在该页面人工完成后，后台仍会核验当前授权，只有当前会话已由服务端确认成功时，点击「继续队列」才记录本次管理员确认并恢复后续项；不会仅凭旧 Microsoft 标签页的短期心跳推进。该确认严格绑定当前 run、session、binding 与本次交接时间，不会创建新票据或扩大密码读取能力。当前项一旦转人工，不再自动填写。手动操作后台前应点击「结束助手 / 切换人工」，然后刷新队列；失败的会话按原有重试流程处理。刷新或关闭后台会使 Microsoft 助手停止，请关闭遗留助手标签页再启动新一轮。

安全边界：管理员 JWT 只从后台 origin 的 localStorage 读取并发送同源管理请求，不写 GM 存储、不进入 Microsoft origin。每个会话仅签发一次最多 90 秒的随机 ticket；Redis 只存 SHA-256 摘要及 sessionId、emailId、adminId、tokenVersion、UA 摘要、有效期。Microsoft 侧一次性领取 ticket 后，获得最长 5 分钟且不超过 device session 有效期的单会话 capability。capability 仅可读取当前会话（包括用于页面流程绑定的非敏感 `client_id`）和领取一次密码，不能列举、指定账号、访问管理 API、轮询 Microsoft token 或进入下一项。

密码只有在脚本识别并核对当前账号的密码输入页后才通过 capability 原子领取一次；领取前、填入前和最终点击前都会重新采集页面，复核运行/会话/标签页绑定、目标邮箱、`client_id` 以及同一输入框、按钮和表单。最终复核与点击之间没有异步等待。密码只在内存中填入目标密码输入框；若点击被 Microsoft 校验阻止而没有导航，脚本仍保留该真实输入框引用，并在随后暂停、超时或成功核验时主动清空。不写 GM、localStorage、剪贴板、URL、状态面板或日志。GM 中的一次性 ticket 领取前即删除，capability 存在 Tampermonkey 的当前标签页隔离状态中；暂停/失联时清除，服务端 TTL 限制始终生效。设备码和邮箱的非密操作证明会在最终复核后、调用受信任 Microsoft 按钮前同步写入一个 GM 单槽，以免新版页面由点击处理器直接导航、且导航中断异步标签页保存；证明严格绑定 runId、sessionId、bindingId、`client_id`、邮箱、标签页绑定时间和到期时间，新会话会覆盖旧值，暂停或成功时删除。该槽绝不包含 ticket、capability 或密码，且只有仍持有当前标签页 capability、通过服务端 `current` 复核并匹配全部绑定的后续页面才能恢复证明。若按钮调用失败，证明立即清除并暂停。密码响应丢失或解析失败后不再次发放，需人工处理或取消并创建新会话。

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
