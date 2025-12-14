## v1.5.6 (2025-12-14)
- **高级/专家日志分开记录**：新增按日滚动的 `app_advanced_YYYY-MM-DD.log` 与 `app_expert_YYYY-MM-DD.log`，用于分别记录高级与专家排错信息。
- **Graph REST 排错日志增强**：Advanced/Expert 级别都会记录 Graph 请求/响应详情（Advanced 记录摘要并截断 body；Expert 记录更完整内容并截断至 50KB）。
- **EWS Trace 路由优化**：EWS 的 Trace 输出会写入当前级别对应的 Advanced/Expert 日志文件；Expert 的 GetItem 响应单独写入 `ews_getitem_responses_expert_YYYY-MM-DD.log`。
- **主界面与工具菜单联动 + Expert 警告**：日志级别在主界面与 Tools 菜单保持同步；选择 Expert 会弹出安全警告确认。

## v1.5.5 (2025-12-14)
- **Graph 会议展开逻辑对齐 EWS**：不填写日期范围时，Graph 与 EWS 一样不展开循环实例（仅扫描主项/单次）；填写开始+结束日期后才使用 `calendarView` 在范围内展开 occurrence/exception。
- **GOID 显示修正**：`MeetingGOID` 对齐为 `iCalUId`（与 EWS 报告一致）；同时最佳努力读取 GlobalObjectId（MAPI 扩展属性）并以更易读形式写入 Details。
- **补齐字段**：Graph Meeting 报告补齐 `UserRole`（按 Organizer/Attendee 推断）与 `ResponseStatus`。
- **循环规则更易读**：Graph 的 `recurrence.pattern` / `recurrence.range` 输出改为易读字符串（不再直接输出原始 JSON）。
- **UI 提示更新**：提示文字改为“填写日期范围才展开循环实例”。

## v1.5.4 (2025-12-14)
- **Graph 会议循环展开**：Meeting 模式使用 `calendarView`（强制要求开始/结束日期），可在时间范围内扫描循环会议的 occurrence/exception。
- **Graph 报告对齐 EWS**：Graph Meeting 报告字段对齐 EWS 报告格式，包含 Organizer/Attendees/Start/End/IsCancelled/ResponseStatus 等，并补充 `iCalUId`、`SeriesMasterId`。
- **GOID 最佳努力**：Graph 尝试通过 MAPI 扩展属性读取 GlobalObjectId（若租户/权限/数据不可用则自动留空，仍可使用 `iCalUId` 追踪）。
- **UI 提示增强**：会议模式下提示 Graph 必须填写日期范围；EWS 不填写日期范围则不展开循环实例。

## v1.5.3 (2025-12-14)
- **UI 更新**：
  - **标题与版本一致性**：标题栏、关于窗口统一显示 v1.5.3。
  - **关于窗口增强**：增加 GitHub 项目链接与头像显示（若缺少 Pillow 则自动降级为仅文本）。
  - **EWS 标签优化**：EWS 配置区域统一显示为 "Exchange On-Premise"。
  - **日期选择器增强**：
    - 点击日期输入框空白处可直接打开日历。
    - 日历支持年份/月分快速选择。
    - 会议模式下，以另一侧已选日期为参考，超出前后两年的日期自动置灰不可选。

## v1.5.2 (2025-12-14)
- **Bug 修复**：
  - **日志显示修复**：修复了重新显示日志区域时，日志目录链接被挤到日志窗口上方的问题。现在日志区域会正确地显示在链接栏上方。

## v1.5.1 (2025-12-14)
- **UI 体验优化**：
  - **智能窗口缩放**：点击“隐藏日志”时，窗口会自动向上收缩，避免留下大片空白区域；点击“显示日志”时自动恢复高度。
  - **布局微调**：确保日志目录和报告链接始终固定在窗口底部，不受日志区域折叠影响。

## v1.5.0 (2025-12-14)
- **安全性增强**：
  - **双重确认机制**：在取消“仅报告”模式（即进入删除模式）时，以及点击“开始清理”时，增加了强制性的二次确认弹窗，防止误操作导致数据丢失。
  - **动态按钮文本**：根据是否勾选“仅报告”，按钮文本会动态显示为“开始扫描 (Start Scan)”或“开始清理 (Start Clean)”，直观提示当前操作风险。
- **日志系统升级**：
  - **按日轮转**：日志文件现在按日期生成（如 `app_2025-12-14.log`），避免单个日志文件过大。
  - **日志分隔**：每次任务开始时，日志中会自动添加分隔符和模式说明（Report Only 或 DELETE），便于区分多次运行记录。
  - **UI 优化**：增加了“显示/隐藏日志”按钮，允许用户折叠日志区域以节省界面空间。

## v1.4.0 (2025-12-14)
- **性能优化**：
  - **多线程执行**：Graph API 和 EWS 清理任务现在使用线程池（ThreadPoolExecutor）并发执行，支持同时处理 10 个用户，显著提升批量处理速度。
  - **线程安全**：引入了文件锁（Lock），确保多线程环境下日志和 CSV 报告写入的安全性。
- **界面优化**：
  - **DPI 适配**：升级了 DPI 感知机制（Per-Monitor V2），在高分屏上显示更加清晰。
  - **布局调整**：增加了窗口最小尺寸限制，优化了字体大小，防止界面元素重叠或过小。

## v1.3.4 (2025-12-13)
- **UI 优化**：
  - 统一了 Graph API 手动配置界面的样式，使其与 LargeAttachmentFinder 保持一致（使用 Grid 布局，标签更清晰）。
  - 确保“环境选择”（Global/China）在手动配置模式下依然可见且生效。

## v1.3.3 (2025-12-13)
- **UI 优化**：
  - 重新调整 Graph API 配置界面布局，确保在手动配置模式下，App ID、Tenant ID 和 Client Secret 输入框正确显示。
  - 优化了环境选择（Global/China）和认证模式选择的布局，使其更加直观。

## v1.3.2 (2025-12-13)
- **Bug 修复**：
  - 修复了 Graph API 手动配置模式下无法选择云环境（Global/China）的问题。现在环境选择对两种模式均可见且生效。

## v1.3.1 (2025-12-13)
- **UI 重构**：
  - 合并“Graph API 设置”和“EWS 设置”为统一的“连接配置”标签页。
  - 将“清理任务”标签页重命名为“任务配置”。
  - 新增连接模式选择（EWS / Graph），选中一种模式时自动禁用另一种模式的配置项。
- **Graph API 增强**：
  - 新增 Graph API 认证模式选择：
    - **自动配置 (Auto)**：使用证书认证，支持一键生成证书和创建 App（保留原有功能）。
    - **手动配置 (Manual)**：支持 Client Credentials Flow，允许手动输入 Client Secret 进行连接（类似 LargeAttachmentFinder）。
- **配置管理**：
  - 配置文件新增 `graph_auth_mode` 和 `client_secret` 字段。

## v1.3.0 (2025-12-13)
- **版本升级**：
  - 升级至 v1.3.0，包含累积的功能增强和 Bug 修复。
- **功能增强**：
  - **日志系统升级**：
    - 新增“工具”菜单，包含“日志配置”子菜单。
    - 支持三种日志级别：
      - **默认 (Default)**：仅记录关键操作日志。
      - **高级 (Advanced)**：额外记录 EWS 请求头和请求体 (不含响应)。
      - **专家 (Expert)**：记录所有 EWS 请求和响应 (含 XML)。开启时会有安全警告。
  - **UI 优化**：日志级别选择区域与菜单栏联动。
  - **Bug 修复**：进一步优化 `RecurrenceDuration` 解析逻辑，增加对 `EndDateRecurrence` 的深度检查，确保能正确读取结束日期。

## v1.2.36 (2025-12-13)
- **功能增强**：
  - **日志系统升级**：
    - 新增“工具”菜单，包含“日志配置”子菜单。
    - 支持三种日志级别：
      - **默认 (Default)**：仅记录关键操作日志。
      - **高级 (Advanced)**：额外记录 EWS 请求头和请求体 (不含响应)。
      - **专家 (Expert)**：记录所有 EWS 请求和响应 (含 XML)。开启时会有安全警告。
  - **UI 优化**：日志级别选择区域与菜单栏联动。
  - **Bug 修复**：进一步优化 `RecurrenceDuration` 解析逻辑，确保能正确读取 `EndDateRecurrence` 中的结束日期。

## v1.2.35 (2025-12-13)
- **功能增强**：
  - 进一步修复 `RecurrenceDuration` 显示“未知”的问题，增强了对 `EndDateRecurrence` 等复杂对象的解析能力。
  - 优化任务完成提示：若仅生成报告，提示“扫描生成报告任务完成”；若执行删除，提示“清理任务已完成”。
  - 增加配置校验：若选择了 Graph 或 EWS 模式但未填写必要配置，会弹出警告并自动跳转到相应的设置标签页。

## v1.2.34 (2025-12-13)
- **功能增强**：
  - 修复了 `RecurrenceDuration` (循环持续时间) 显示为“未知”的问题。现在能正确识别“无限期”、“结束于”或“共N次”。
  - 优化了 `IsEndless` 的判断逻辑，使其更准确。
- **说明**：
  - “周首日” (FirstDayOfWeek) 指的是循环模式中定义的每周开始的第一天（例如周一或周日），这会影响跨周期的计算。

## v1.2.33 (2025-12-13)
- **严重 Bug 修复**：
  - 修复了 v1.2.32 中 `EwsTraceAdapter` 缺少返回值导致 `AttributeError: 'NoneType' object has no attribute 'elapsed'` 的崩溃问题。

## v1.2.32 (2025-12-13)
- **EWS 响应日志增强**：
  - 改为**实时写入** `ews_getitem_responses.log`，不再使用内存缓冲，彻底解决日志丢失问题。
  - 启动时自动检测并记录响应日志文件路径，若无权限写入会提示警告。
  - 仅捕获 XML 或 Text 类型的响应，避免二进制数据干扰。
- **Bug 修复**：
  - 修复了 v1.2.30/31 中可能因 finally 块执行顺序导致的日志未写入问题。

## v1.2.31 (2025-12-13)
- **Pattern 信息中文化**：
  - `RecurrencePattern` 列现在显示中文类型（如"按周"、"按天"、"按月(相对)"）。
  - `PatternDetails` 列内容汉化（如"星期=周一, 周三"、"间隔=1"、"日期=15日"）。
  - `RecurrenceDuration` 列内容汉化（"无限期"、"结束于: 2025-12-31"、"共 10 次"）。
- **Bug 修复**：
  - 修复了 v1.2.29 中引入的 `UnboundLocalError: cannot access local variable 'os'` 错误（v1.2.30 已修复，此版本包含该修复）。

## v1.2.29 (2025-12-13)
- **Pattern 详细信息列**：
  - 添加 `PatternDetails` 列，显示 Pattern 的完整属性（如"WeeklyPattern: Days=Mon, Wed, Fri, Interval=1"、"DailyPattern: Interval=2"）。
  - 添加 `RecurrenceDuration` 列，显示循环持续时长（"Endless=True"、"EndDate: 2025-12-31"、"Occurrences: 10"）。
  - 添加 `IsEndless` 列，仅当 `Type=RecurringMaster` 且循环无限期时写 "True"，其他情况写 "N/A"。
- **新增辅助函数**：
  - `get_pattern_details(pattern_obj)`：提取 Pattern 对象的 Interval、Days、Month、DayOfMonth 等属性，生成易读的详细字符串。
  - `get_recurrence_duration(recurrence_obj)`：提取循环持续方式（EndDate/NoEnd/MaxOccurrences）。
  - `is_endless_recurring(item_type, recurrence_obj)`：判断是否为无限期的 RecurringMaster。

## v1.2.28 (2025-12-13)
- **EWS 响应日志修复**：TraceAdapter 缓冲响应，在任务 finally 时写入 `ews_getitem_responses.log`，确保文件始终生成（避免权限/路径问题）。
- **Pattern 计算改进**：
  - 添加 Pattern 缓存机制，相同 UID 的实例共享 Pattern。
  - Instance 查找 master 时记录日志，便于诊断（如"通过 RecurringMasterId 获得主项"、"通过 UID 找到主项"）。
  - Pattern 计算前检查缓存，减少重复查询。

## v1.2.27 (2025-12-13)
- **Master/Instance 判定**：
    - 有 `recurrence` ⇒ `RecurringMaster`
    - 有 `original_start` 或 `recurrence_id` ⇒ `Instance`
    - 其他走 `guess_calendar_item_type` 兜底。
- **Occurrence/Exception 细分（仅 Instance）**：
    - `start != original_start` ⇒ `Exception`（确定）
    - 否则尝试拿主项（先 `recurring_master_id`，再 `uid`）：对比 `subject/location/attendees`，不同 ⇒ `Exception`，否则 `Occurrence`；拿不到主项 ⇒ `Instance-Unknown`。
- **Pattern 计算**：仅在 `RecurringMaster` 且存在 `recurrence` 时记录，使用 `pattern.__class__.__name__`；实例不计算，留空/N/A。

## v1.2.25 (2025-12-13)
- **类型判断兜底推断**：新增 `guess_calendar_item_type`，优先使用服务器返回的 CalendarItemType，缺失时按 recurrence/recurrence_id/original_start/recurring_master_id/is_recurring 逐级推断。
- **日志诊断增强**：Advanced 日志输出 GuessType、RawType、UID、IsRecurring、RecMasterId、RecurrenceId、OriginalStart、HasRecurrence、PyType、ClassName、ItemClass，便于快速定位非标准项。
- **过滤与写行统一使用推断类型**：scope 过滤和行输出都使用 guess 结果，Series 支持 RecurringMaster/Occurrence/Exception。
- **Pattern 填充逻辑简化**：仅在推断为 RecurringMaster 且存在 recurrence 时记录 pattern（使用 pattern 类名），其他情况留空避免误判。

## v1.2.24 (2025-12-13)
- **GetItem 响应日志改进**:
    - TraceAdapter 现在完整捕获 GetItem 操作的响应 XML
    - GetItem 响应同时被写入单独文件: `%USERPROFILE%\Documents\UniversalEmailCleaner\ews_getitem_responses.log`
    - 便于离线查看 `<t:CalendarItemType>` 服务端返回值，排查 Type=Unknown 问题

## v1.2.23 (2025-12-13)
- **关键诊断改进 (Critical Diagnostics)**:
    - 增强 Advanced 日志：添加 `ItemClass`、`PyType`、`ClassName` 输出，用于识别是否为 IPM.Appointment 或其他非标准项类（如 IPM.Schedule.Meeting.Request）。
    - 改进 TraceAdapter 响应记录：确保 GetItem 响应 XML 被完整捕获（截断至 50KB），包含服务端实际的 `<t:CalendarItemType>` 值。
- **类型检测推断逻辑 (Type Detection Inference)**:
    - 若 `calendar_item_type` 为空，使用更稳定的推断规则：
      - 有 `recurrence` ⇒ RecurringMaster
      - 有 `recurring_master_id` ⇒ Occurrence
      - `is_recurring == False` ⇒ Single
      - 其他 ⇒ Unknown
    - 避免硬卡 `calendar_item_type` 导致整体失败。
- **RecurringMasterId 兼容性**:
    - 处理 `recurring_master_id` 为 ItemId 对象的情况，提取其 `.id` 和 `.changekey` 后再调用 `account.calendar.get()`。

## v1.2.21 (2025-12-13)
- 修复 CalendarView `.only()` 导致的 `InvalidField` 错误，改为使用纯 `view()` 并在后续通过 `GetItem(additional_fields)` 富化每个项，确保 `calendar_item_type`、`uid`、`recurring_master_id`、`recurrence` 等字段完整。

## v1.2.20 (2025-12-13)
- **修复 (Fixes)**:
    - CalendarView.view() 使用 `.only()` 方法而非 `additional_fields` 参数指定返回字段。
    - 显式请求 `calendar_item_type`、`uid`、`recurring_master_id`、`recurrence` 等关键字段。
    - 修复 "FolderCollection.view() got an unexpected keyword argument" 错误。

## v1.2.18 (2025-12-13)
- **修复 (Fixes)**:
    - CalendarView 使用 `additional_fields` 参数而非 `.only()` 确保返回 `calendar_item_type`、`uid`、`recurring_master_id`、`recurrence` 等关键字段。
    - 修复 Type 显示为 Unknown 和 RecurrencePattern 为空的问题。

## v1.2.17 (2025-12-13)
- **修复 (Fixes)**:
    - `GetItem` 显式请求 `RecurringMasterId`、`CalendarItemType`、`UID`、`Recurrence`、`IsRecurring`，避免日历视图返回的浅项缺失关键字段。
    - UID 兜底查找主项时，若按主题筛选不到结果，会回退到全量扫描，保证能拿到循环规则。
    - `CalendarView` 查询时使用 `.only(*extra_fields)` 显式请求关键字段，提升性能并确保字段完整性。

## v1.2.16 (2025-12-13)
- **修复 (Fixes)**:
    - GetItem 时显式请求 `RecurringMasterId`、`CalendarItemType`、`UID`、`Recurrence`、`IsRecurring`，确保实例能跳回主项获取循环规则。
    - 若缺少 `RecurringMasterId`，改为基于 UID 的客户端匹配主项（`calendar_item_type == RecurringMaster`）以兜底获取 `RecurrencePattern`。

## v1.2.14 (2025-12-13)
- **修复 (Fixes)**:
    - 移除对 `only()` 的依赖以避免在 CalendarView/文件夹上下文中出现 `Unknown field path` 错误。
    - 改为安全的属性访问（`getattr`）读取 `calendar_item_type`、`uid`、`recurring_master_id`、`recurrence`。

# 版本迭代记录 (Version History)

## v1.2.13 (2025-12-13)
- **修复 (Fixes)**:
    - 修复字段名错误：`only()` 中应使用 `id` 而不是 `item_id`。


## v1.2.12 (2025-12-13)
- **重大改进 (Major Improvements)**:
    - **彻底重写会议类型识别逻辑**：直接使用 EWS 返回的 `CalendarItemType`（Single/RecurringMaster/Occurrence/Exception），不再推断。
    - **使用 RecurringMasterId 获取循环规则**：对于 Occurrence/Exception 类型，通过 `recurring_master_id` 字段直接获取主项，更可靠。
    - **补充关键字段**：在查询时显式指定 `uid`、`recurring_master_id`、`recurrence` 等关键字段。
    - 消除了之前版本中的 "Unknown" 类型和大量不必要的批量查询。
    - 显著提升准确性和性能。

## v1.2.11 (2025-12-13)
- **优化 (Improvements)**:
    - 优化会议类型 (Type) 识别逻辑，在 CalendarView 模式下更准确地判断 Occurrence/RecurringMaster。
    - 实现循环规则批量预加载机制，避免为每个 Occurrence 重复查询主项。
    - 显著提升处理大量循环会议实例的性能。
    - Occurrence 类型的会议现在也能正确显示 RecurrencePattern。

## v1.2.10 (2025-12-12)
- **修复 (Fixes)**:
    - 修复获取会议循环规则时的字段过滤错误 (`Unknown field path 'calendar_item_type'`)。
    - 现在只按主题 (Subject) 过滤，然后在客户端检查 UID 和 `calendar_item_type`。
    - 提高了获取循环会议主项 (RecurringMaster) 的可靠性。

## v1.2.9 (2025-12-12)
- **修复 (Fixes)**:
    - 修复与 NTLM 认证库 (`requests_ntlm`) 的兼容性问题。
    - `EwsTraceAdapter` 现在使用 `response._content` 而不是 `response.content` 来避免干扰流式响应和 NTLM 认证流程。
    - 确保 `response.raw` 对象对 NTLM 认证库保持可用。

## v1.2.8 (2025-12-12)
- **修复 (Fixes)**:
    - 进一步修复 EWS 高级日志崩溃问题 (`'NoneType' object has no attribute 'raw'`)。
    - 增强 `EwsTraceAdapter` 的兼容性，使用 `*args` 和 `**kwargs` 传递参数，防止因参数签名不匹配导致的错误。
    - 增加详细的 Traceback 日志输出，便于排查深层错误。

## v1.2.7 (2025-12-12)
- **修复 (Fixes)**:
    - 修复 EWS 高级日志在处理流式响应 (Streamed Response) 或非文本请求体时可能导致的崩溃问题 (`'NoneType' object has no attribute 'raw'`).
    - 优化日志记录逻辑，避免消耗请求/响应流。

## v1.2.6 (2025-12-12)
- **日志增强 (Logging Enhancement)**:
    - 实现了标准的 EWS XML Trace 日志格式。
    - 高级日志模式下，现在会输出包含 `<Trace Tag="EwsRequest" ...>` 和 `<Trace Tag="EwsRequestHttpHeaders" ...>` 的完整 XML 请求与响应记录，格式与 EWSEditor/MFCMAPI 兼容。

## v1.2.5 (2025-12-12)
- **日志增强 (Logging Fixes)**:
    - 修复 EWS 高级日志不显示 HTTP Headers 的问题。
    - 现在同时启用 `exchangelib.protocol` (XML Body) 和 `urllib3` (HTTP Headers/Connection) 的调试日志。

## v1.2.4 (2025-12-12)
- **修复 (Fixes)**:
    - 修复 EWS 无法通过 `uid` 字段筛选会议主项的问题 (`Error: EWS does not support filtering on field 'uid'`)。
    - 改为通过 `Subject` (主题) 和 `CalendarItemType` (RecurringMaster) 查找主项，并在客户端匹配 UID。

## v1.2.3 (2025-12-12)
- **日志增强 (Logging Enhancement)**:
    - 增强 "高级日志 (Advanced Log)" 模式：
        - **EWS**: 启用 `exchangelib` 协议层调试日志，记录完整的 SOAP XML 请求头 (Headers) 和响应体 (Body)。
        - **Graph API**: 记录完整的 HTTP 请求 (GET/DELETE) 的 URL、Headers、Params 以及响应的 Headers 和 Body。
    - 优化日志写入性能，避免大量调试数据导致界面卡顿。

## v1.2.2 (2025-12-12)
- **报表增强 (Report Enhancement)**:
    - 会议模式导出报表新增详细字段：
        - `Type`: 会议类型 (Single, Occurrence, RecurringMaster)
        - `MeetingGOID`: 会议唯一标识符 (UID)
        - `CleanGOID`: 清理后的 GOID
        - `Organizer`: 会议组织者
        - `Attendees`: 与会者列表
        - `Start`/`End`: 会议开始/结束时间
        - `UserRole`: 用户角色 (Organizer/Attendee)
        - `IsCancelled`: 是否已取消
        - `ResponseStatus`: 响应状态 (Accepted, Tentative, Declined, Unknown)
        - `RecurrencePattern`: 循环规则 (仅针对 RecurringMaster)
- **修复 (Fixes)**:
    - 修复 `CalendarView` 筛选限制导致的错误 (ErrorInvalidRestriction)。
    - 修复 `CalendarItem` 属性访问错误。

## v1.2.0 (2025-12-12)
- 优化会议清理逻辑：
    - 引入 EWS `CalendarView` (日历视图) 支持。
    - 支持搜索并删除循环会议 (Recurring Meetings) 在指定时间范围内的单次实例 (Occurrences)。
    - 解决了循环会议主项 (Series Master) 在搜索范围之外，但实例在范围内无法被发现的问题。
- 优化日期输入处理：自动修正多种日期格式 (YYYY/MM/DD, YYYYMMDD)。
- 修复时区处理错误 (`EWSTimeZone` localize issue)。

## v1.1.0 (2025-12-12)
- 新增功能：支持删除会议/日历事件 (Meeting/Calendar Events)
- 新增功能：会议筛选条件
    - 组织者 (Organizer)
    - 会议时间范围 (Start/End Time)
    - 会议类型范围：单次会议 (Single) 或 系列会议主项 (Series Master)
    - 仅删除已取消的会议 (Is Cancelled)
- 界面更新：新增 "清理目标" 选项 (邮件/会议) 及相关会议设置

## v1.0.0 (2025-12-12)
- 初始版本发布
- 支持 Microsoft Graph API 和 Exchange Web Services (EWS) 双模式
- 支持通过 CSV 批量导入目标用户
- 支持多种搜索条件：Message ID, 主题, 发件人, 日期范围, 正文关键字
- 支持 "仅报告" 模式和 "删除" 模式
- 支持全球版 (Global) 和世纪互联 (China) 环境
- 支持 EWS 模拟 (Impersonation) 和代理 (Delegate) 权限
- 包含详细的运行日志和 CSV 报告生成
- 全中文图形化界面
