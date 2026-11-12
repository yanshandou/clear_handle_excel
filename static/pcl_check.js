// PCL 单体测试检查台（static/UIPCL.html，网址 /pcl）：22 项检查清单 + 源码静态校验 + 覆盖率 + 每一项的测试结果报告。
// 结果按「检查对象」保存在程序数据库（core/pcl_check.py）；源码规则在 static/pcl_rules.js，浏览器里跑。
(function () {
'use strict';
const {SOURCE_RULES, SOURCE_STATUS, SOURCE_TYPES, SOURCE_LIMIT, sourceType, checkSources,
  coveragePercent, coverageRanges, coverageFinding} = window.PclRules;

// 已按用户提供的原文逐项人工整理；代码片段仅是示意定位，不代表项目源码。
const ITEMS = [
{ id:'T01', title:'DB 保存后重新加载并回显全部字段', tags:['twig'], jp:'DB登録後に画面を再表示（dataLoad）(全項目の再表示が出来る事)していること', meaning:'数据库新增或更新成功以后，必须调用画面再加载逻辑（通常为 dataLoad），从最新数据重新填充整个画面，并确保所有字段都正确回显。不能只更新部分 UI 就视为通过。',steps:['先填写包含文本、下拉框、日期、复选框等不同类型字段的测试数据。','执行新增或更新，确认数据库保存成功。','确认调用 dataLoad 或等效的重新获取数据逻辑；重新显示后逐项比对保存值。','尝试页面重新打开，核对是否仍与 DB 一致。'],keys:['dataLoad','save','insert','update','setValues','reload'],code:'// 保存成功的回调附近重点查找\nawait saveData(payload);\ndataLoad(); // 或项目规定的等效刷新逻辑',note:'重点检查隐藏字段、下拉选中值、编辑状态、异步请求顺序；这里只是定位示例，不要求代码逐字一致。'},
{ id:'T02',title:'初始化函数必须彻底复原画面状态',tags:['twig'],jp:'初期化用functionを流すことで、値や表示が必ず初期化される作りになっていること',meaning:'执行画面的初始化函数时，输入值、组件显示/隐藏、可用/不可用、选中状态等必须恢复至规格规定的初始状态，不应残留上一次操作的数据。',steps:['先手动修改各类值，切换标签、选择行并改变控件显示状态。','调用初始化函数，或执行会触发初始化的操作。','核对输入内容、隐藏显示、按钮状态、列表选择和校验提示均恢复初始状态。'],keys:['init','initialize','reset','clear','setValues','show(','hide(','enable(','disable('],code:'function initForm() {\n  $$ ("form").clear();\n  // 检查：选择状态、隐藏控件和按钮状态是否也复原\n}',note:'函数名不一定是 init。应根据设计书检查“应当保留”的默认值，不要把清空所有内容误认为正确初始化。'},
{ id:'T03',title:'数值列必须使用整数排序',tags:['twig','webix'],jp:'一覧数値列のソートが「sort:"int"」となっていること',meaning:'Webix 数据表中的数值列需按整数大小排序，而非当作字符串按字典序排序。比如 2 应排在 10 之前。',steps:['定位目标 datatable 的 columns 定义。','核对所有数值列使用 sort:"int"（按项目约定）。','用 2、10、100 等数据测试升序/降序，确认结果。'],keys:['datatable','columns','sort: "int"','sort:"int"'],code:'{ id: "quantity", header: "件数", sort: "int" }',note:'要区分数字列和字符串编码列，带前导零的编号不一定应按整数处理。'},
{ id:'T04',title:'可拖拽排序的列表禁止列排序',tags:['twig','webix'],jp:'ドラッグドロップが可能な一覧は、列ソートが不可であること',meaning:'支持拖动调整行顺序的列表，不允许再通过列标题进行排序，以免“拖拽顺序”和“列排序”产生冲突。',steps:['找到存在行拖拽功能的 datatable。','检查所有列是否没有开启用户可用的排序。','运行画面：拖拽行后点击列标题，确认没有触发排序或破坏行序。'],keys:['drag','sort','columns','onAfterDrop','datatable'],code:'// 对可拖动列表检查 columns 是否配置了 sort\n{ view: "datatable", drag: "order", columns: [\n  { id: "name", header: "名称" }\n] }',note:'程序调用的初始化排序、数据查询 ORDER BY 与“用户点击列标题排序”需分开判断。'},
{ id:'T05',title:'保存行选中状态的列表要有隐藏 id 列',tags:['php','twig'],jp:'一覧にて行選択状態を保存する場合に、隠し列「id」があること',meaning:'如果列表需要在刷新或重新取得数据后恢复之前选中的行，要保留能唯一定位该行的 id（通常为隐藏列），并核对 PHP 传入的字段。',steps:['确认该列表确实需要保存/恢复选中行。','检查 Twig 的 datatable 配置是否存在隐藏 id 列，以及 PHP 数据有无对应唯一值。','选择某行→刷新/重新绑定数据→核对仍选中同一条记录。'],keys:['id','hidden: true','select','getSelectedId','datatable'],code:'columns: [\n  { id: "id", hidden: true },\n  { id: "name", header: "名称" }\n]',note:'不要仅凭 Webix 内部自动生成的 item.id 就判定满足原表要求；需确认项目所指的隐藏 id 列。'},
{ id:'T06',title:'使用代码主数据时记录 Code Master ID',tags:['docs'],jp:'コードマスタを使用する場合は、概要書の備考欄にコードマスタIDの記載があること',meaning:'画面使用了代码主数据（コードマスタ）时，必须在对应概要书的备注栏写清所用的代码主数据 ID，方便追踪选项来源和维护。',steps:['定位画面中通过代码主数据生成的下拉选项、显示文本等。','从实现或接口定义中确认对应代码主数据 ID。','打开概要书，核对目标条目的「備考」栏是否记录该 ID。'],keys:['コードマスタ','コードマスタID','備考','codeMaster','masterId'],code:'// 查找代码主数据读取的位置\n// 然后对照「概要書 → 備考欄 → コードマスタID」',note:'这是文档一致性检查，不是单纯在源码中存在 ID 就通过。'},
{ id:'T07',title:'Debug 模式不得出现重复 ID 警告',tags:['twig'],jp:'Debugモードで画面を表示する。その際に、IDが重複している旨のメッセージが左下に出ないこと',meaning:'在 Debug 模式下打开目标画面，确认左下角不会弹出 ID 重复相关的警告；需要检查动态组件创建时使用的 ID 是否唯一。',steps:['启用项目规定的 Debug 模式。','打开目标画面并执行可能创建子画面/弹窗的操作。','观察左下角是否出现重复 ID 提示，同时检查浏览器控制台信息。'],keys:['id:','$$(','webix.ui','debug','destructor','onShow'],code:'// 重点排查重复创建时固定不变的组件 id\n{ view: "text", id: "customer_name" }',note:'页面首次显示没有警告并不够，重复打开弹窗、切换页面和回到主页面后也要检查。'},
{ id:'T08',title:'PHP foreach 不直接整体赋值 value',tags:['php'],jp:'foreach内にて、valueをそのまま変数に格納していないこと（項目を一個一個明記すること）',meaning:'在 foreach 处理 DB 或接口返回数组时，不要不加筛选地把整条 value 赋给输出变量；应逐个明确指定需要的字段，提高接口字段可控性。',steps:['全局检索 foreach，重点查看响应数据加工、列表数据组装位置。','检查循环体是否把整行 $value 直接赋值或追加给结果数组。','确认按照规定逐个列明输出字段，不带入无用或敏感字段。'],keys:['foreach','=> $value','$result[] = $value','array_map','return $rows'],code:'// 避免：$result[] = $value;\nforeach ($rows as $value) {\n    $result[] = [\n        \'id\' => $value[\'id\'],\n        \'name\' => $value[\'name\']\n    ];\n}',note:'本检查针对“直接把 value 整体塞入结果”；foreach ($rows as $value) 这种循环变量本身并非违规。'},
{ id:'T09',title:'form 内非 datatable 控件禁止动态宽度设置',tags:['webix'],jp:'form内のdatatableを除く項目にwidthが変動する設定をしていないこと(minWidth,maxWidth等)',meaning:'Webix form 中，除 datatable 外，不要对表单项配置会让宽度动态变化的属性，如 minWidth、maxWidth 等；按项目要求固定控制布局。',steps:['找到 view:"form" 及其 elements/rows/cols 中的控件。','排除 datatable 后检查 minWidth、maxWidth、自动伸缩相关配置。','调整窗口尺寸，确认控件布局没有不符合规格的变化。'],keys:['view: "form"','minWidth','maxWidth','width','elements','responsive'],code:'// form 内除 datatable 外重点检索\n{ view: "text", name: "title", minWidth: 120 }\n// ↑ 属于需要核对/整改的写法',note:'原要求明确排除了 datatable；不要误删 datatable 合法的宽度控制。'},
{ id:'T10',title:'label 等组件禁止 autowidth:true',tags:['webix'],jp:'label等のautowidth:trueは使用禁止（このプロパティは英数字にのみ作用する）',meaning:'label 等组件不应使用 autowidth:true（或项目对应的 autoWidth 写法），因为该属性按原规范只对英数字有适用性，日文/中文布局可能异常。',steps:['全局检索 autowidth / autoWidth 配置。','重点检查 label 和标签相关组件。','确认没有违规使用；必要时以明确宽度或项目约定的布局方式替代。'],keys:['autowidth','autoWidth','view: "label"','labelWidth'],code:'// 禁止的示意写法\n{ view: "label", label: "利用者名", autowidth: true }',note:'项目规范使用的是 autowidth:true；注意大小写差异，避免遗漏实际源码中的写法。'},
{ id:'T11',title:'可拖动的 datatable 必须使用 drag:"order"',tags:['webix'],jp:'ドラッグ&ドロップを可とするdatatableは「drag:"order"」となっていること（「drag:true」は使用しない）',meaning:'如果 Webix datatable 允许行拖拽，drag 配置必须是 "order"，不能使用 drag:true；这样语义是调整行顺序而非任意拖放。',steps:['全局搜索 datatable 的 drag 配置。','逐一核对有行拖拽需求的列表使用 drag:"order"。','实测拖拽顺序，并核对保存/重新加载后顺序是否符合设计要求。'],keys:['drag: "order"','drag:"order"','drag: true','datatable'],code:'{ view: "datatable", drag: "order" }\n// 避免：{ view: "datatable", drag: true }',note:'同时和第 04 项一起确认列排序禁用。'},
{ id:'T12',title:'最大字符数：保存并回显不能截断',tags:['webix'],jp:'最大文字数チェックは、以下の手順で行うこと。\n①確認する項目に最大文字数で入力する。\n②登録処理後、画面を再表示し、文字切れしていないこと。（登録時と完全一致）',meaning:'对存在最大输入长度限制的字段，输入恰好达到上限的文本，保存后重新显示画面，确保数据一字不差，没有因为 UI、DB、编码或接口而截断。',steps:['根据设计书查明字段最大允许字符数 N。','输入恰好 N 个字符的文本并保存。','重新打开/调用 dataLoad，逐字确认与输入内容完全一致。','对项目适用的日文、中文、emoji 等多字节字符另行取样。'],keys:['maxlength','maxLength','max:','varchar','mb_strlen','dataLoad'],code:'// 输入长度限制和 DB 字段长度应对齐\n{ view: "text", name: "title", attributes: { maxlength: 20 } }',note:'“字符数”“UTF-8 字节数”“数据库字段长度”的度量方式可能不同。不要只看没有报错，还要比对回显的完整内容。'},
{ id:'T13',title:'最小字符数：空值和下限输入要正常处理',tags:['webix'],jp:'最小文字数チェックは、以下の手順で行うこと。\n①確認する項目に空、必須項目は最小文字数で入力する。\n②登録処理後、画面を再表示し、エラーが発生せず、登録時と完全一致すること。',meaning:'测试可选字段为空，以及必填字段恰好达到允许的最小字符数。完成登记后再次加载画面，确认没有异常且保存值完全一致。',steps:['根据设计书区分必填字段和非必填字段。','非必填字段输入空值；必填字段输入恰好满足最小长度的内容。','保存并重新显示，确认没有错误并且字段值与登记时一致。'],keys:['minLength','minlength','required','validate','rules','NOT NULL','dataLoad'],code:'// 结合设计书确认：必填 / 最小长度 / 空值处理\n{ view: "text", required: true, name: "code" }',note:'不要把“必填字段输入空值”作为此条正向通过用例；原文要求其输入最小字符数。'},
{ id:'T14',title:'Twig 统一使用双引号（标签内除外）',tags:['twig'],jp:'「\'」を使っていないこと。(「\"」で統一すること) ※タグ内は除く',meaning:'按照项目约定，Twig 模板中的相关代码字符串统一使用双引号，不使用单引号；原要求特别说明“标签内除外”。',steps:['在变更过的 Twig 文件中查找单引号字符。','排除 HTML 标签内部及规范列明的例外。','检查 JS 字符串、组件配置等是否需要改成双引号。'],keys:['.twig','\'','"','view:','id:'],code:'// Twig / JS 配置示例\n{ view: "text", id: "user_name" }\n// HTML 标签内的引号按原文为例外',note:'自动替换所有单引号容易破坏 HTML 属性、Twig 语法或字符串转义；务必区分上下文。'},
{ id:'T15',title:'PHP 统一使用单引号',tags:['php'],jp:'「\"」を使っていないこと。(「\'」で統一すること)',meaning:'根据项目 PHP 代码风格，字符串优先且按检查项要求统一使用单引号，避免新增双引号字符串。',steps:['在本次修改过的 PHP 文件中定位双引号字符串。','确认是否可用单引号替代，处理单引号转义和不需要插值的字符串。','复核 PHP 语法和原本的字符串实际内容没有改变。'],keys:['.php','"','\'','sprintf','foreach'],code:'// 按项目规范\n$message = \'登録完了\';\n// 不使用：$message = "登録完了";',note:'PHP 单/双引号的变量插值和转义规则不同，不能盲目批量替换导致行为变化。'},
{ id:'T16',title:'变更必须落在 Redmine 票据指定环境',tags:['php','twig','release'],jp:'Redmineチケットのリリースバージョンに対して、正しい環境(dev,mainte)に変更が適用されていること',meaning:'核对 Redmine Ticket 的 Release Version 与代码实际提交/部署的分支或环境一致，确保变更进入正确的 dev 或 mainte，既不遗漏也不误入其他维护线。',steps:['查看目标 Redmine Ticket 的リリースバージョン。','确认相应 dev/mainte 分支、提交和差异文件。','在有权限的前提下核对实际环境应用状态，并记录证据。'],keys:['Redmine','リリースバージョン','dev','mainte','git branch','git log','git diff'],code:'git branch --show-current\ngit log --oneline -n 5\ngit diff origin/dev -- path/to/file',note:'Git 分支上存在修改不等于已经部署。需要明确当前代码、合并状态与目标环境的关系。'},
{ id:'T17',title:'include 子画面 ID 时必须传 sub_torihiki_id',tags:['twig'],jp:'子画面IDをincludeする時、必ずsub_torihiki_idを設定すること',meaning:'Twig 中 include 子画面 ID 或包含子画面模板时，必须同时设置 sub_torihiki_id，避免子画面业务上下文缺失。',steps:['检索 include 及相关子画面嵌入位置。','确认每处子画面 ID 的 include 都传入/设置 sub_torihiki_id。','实测子画面打开、保存和切换时该值正确。'],keys:['include','sub_torihiki_id','child','子画面ID'],code:'{% include "child.twig" with {\n  "sub_torihiki_id": sub_torihiki_id\n} %}',note:'上面是 Twig 风格的示意，不一定是项目既有的 include 格式；以实际模板接口为准。'},
{ id:'T18',title:'validate 中必须执行表单校验并返回 false',tags:['twig'],jp:'validate functionの中に、必ずif(!$$("form").validate()){\n  webix.message(CONST.MSG.VL10001);\n  return false;\n}を入れること。',meaning:'目标画面的 validate function 必须调用表单的 validate()；如果校验失败，显示统一提示 CONST.MSG.VL10001，并立即 return false，阻止后续登记。',steps:['找到当前画面的 validate function。','确认存在项目规定的表单校验、统一错误消息和 return false。','留空必填项或输入不合法值，点击登记，确认显示消息且请求未继续执行。'],keys:['function validate','$("form")','$$("form")','validate()','CONST.MSG.VL10001','return false'],code:'if (!$$("form").validate()) {\n  webix.message(CONST.MSG.VL10001);\n  return false;\n}',note:'注意 form 的实际 ID 是否确实为 "form"；检查提示信息和早退分支均能执行。'},
{ id:'T19',title:'跨 Tab 取数据需设置第二参数 true',tags:['twig'],jp:'tabviewについて、選択中のタブ以外のデータも取りたい場合は、getItemやgetItemAllの第二引数にtrueをセットしていること。',meaning:'在 tabview 中，如果业务逻辑需要读取当前未选中 Tab 的数据，调用 getItem 或 getItemAll 时必须把第二个参数设为 true，以便包含非活动标签页。',steps:['检索各 tabview 中的 getItem / getItemAll 调用。','判断是否需要取得非选中 Tab 的数据。','对需要跨 Tab 的调用确认第二实参为 true。','在其他 Tab 填入数据后切回当前 Tab 执行保存/取值并核对。'],keys:['tabview','getItem(','getItemAll(','true'],code:'// 需要读取未激活 Tab 的数据时\ngetItem(itemId, true);\ngetItemAll(tabId, true);',note:'getItem/getItemAll 可能是项目封装 API，而不是 Webix 原生通用接口；应查看项目定义确认参数含义。'},
{ id:'T20',title:'列表图标列必须指定排序项',tags:['twig','webix'],jp:'一覧アイコン列のソート項目が指定されていること',meaning:'列表中以图标呈现状态的列，需要显式指定对应的排序项或排序逻辑，确保排序依据为图标代表的业务值，而不是单纯 HTML 或显示模板。',steps:['定位 datatable 中模板输出图标的列。','确认其配置有明确的排序字段/排序规则。','用不同图标状态的数据进行升序降序测试。'],keys:['icon','template','sort','columns','datatable'],code:'{\n  id: "status",\n  template: statusIconTemplate,\n  sort: "string" // 需对应真实业务字段/规则\n}',note:'原文只要求指定“排序项”，不是所有项目都统一 sort:"string"；若此列表允许行拖拽，还需一并遵守第 04 项。'},
{ id:'T21',title:'检查代码覆盖率 100% 与分支遗漏',tags:['quality'],jp:'カバレージ100%を確認する事。AIツール（コーディングの読み合わせ、またはAIツールを活用し、分岐に漏れがないことを確認する）\n※消化は机上または打鍵で確認する',meaning:'需要确认测试覆盖率达到项目要求的 100%，同时通过代码走查或 AI 辅助检查遗漏的业务分支；每个检查点可以通过机上评审或实际操作确认。',steps:['确认本次变更的受测范围及项目覆盖率计算口径。','运行适用的测试并查看覆盖率报告（如 PHPUnit/Xdebug）；确认要求范围达到 100%。','逐条走查 if/elseif、异常、边界、空值等分支，必要时使用 AI 辅助审阅。','记录人工走查或实际按键测试的证据和未覆盖的原因。'],keys:['PHPUnit','Xdebug','coverage','if','elseif','else','try','catch'],code:'// PHPUnit 常见示意，具体命令以项目配置为准\nvendor/bin/phpunit --coverage-html coverage',note:'行覆盖率 100% 不自动等于分支覆盖率 100%；不能只让 AI 口头判断就标记通过。需用可核验的测试/走查证据。'},
{ id:'T22',title:'DB 查询只获取实际使用的字段',tags:['quality','php'],jp:'DB取得時、使用するカラムのみ取得していること',meaning:'数据库查询应只 SELECT 真正需要的列，避免 SELECT * 或无用字段，减少数据传输、暴露和后续耦合。',steps:['搜索本次变更涉及的 SQL 和查询构造器。','核对 SELECT 列是否与 PHP/Twig 真正使用的字段一致。','剔除无用字段，检查关联查询和别名字段不会缺失。','运行功能测试确认删除无用列后仍能正确输出。'],keys:['SELECT *','select(','->select(','PDO','query(','fetchAll'],code:'-- 不建议：SELECT * FROM users\nSELECT id, display_name, status\nFROM users\nWHERE deleted = 0;',note:'有些查询需要关联字段作过滤或排序，未输出不代表无用；应区分读取列和 WHERE/ORDER BY 所依赖的列。'}
];
const CATS=[{id:'all',name:'全部项目'},{id:'twig',name:'Twig / 画面'},{id:'php',name:'PHP / 后端'},{id:'webix',name:'Webix / 组件'},{id:'docs',name:'概要书 / 文档'},{id:'release',name:'发布环境'},{id:'quality',name:'质量 / DB'}];
const LABELS={twig:'TWIG',php:'PHP',webix:'WEBIX',docs:'DOCS',release:'RELEASE',quality:'QUALITY'};

const AUTO_ITEMS = [...Object.keys(SOURCE_RULES), 'T21'].sort();
const RUN_LABELS = {...SOURCE_STATUS, none:'没有相关写法'};
const VERDICTS = {ok:{label:'OK', mark:'✓'}, ng:{label:'NG', mark:'✗'}, na:{label:'不适用', mark:'—'}, '':{label:'待判定', mark:''}};
const THEME_KEY = 'pcl-qa-theme';
const TARGET_KEY = 'pcl-qa-target';
const TESTER_KEY = 'pcl-qa-tester';
const LEGACY_KEY = 'pcl-qa-checklist-v1';  // 旧版页面存在浏览器里的进度
const LEGACY_ASKED_KEY = 'pcl-qa-legacy-asked';
const IMAGE_LIMIT = 8 * 1024 * 1024;
const $ = id => document.getElementById(id);

let targets = [], target = null, results = {}, coverage = null;
let activeCategory = 'all', activeStatus = 'all', query = '', selected = 'T01';
let sourceFiles = [], sourceWarnings = new Set(), sourceSig = '';
// 不是每一项都要测：selection = 这个对象勾选了哪些项目；runs = 每项最近一次点「测试」的结果
let selection = new Set(ITEMS.map(x => x.id)), runs = {};
let nextSourceNumber = 1, coverageOnlyCards = false, shownNoteFor = '';

function read(key) { try { return localStorage.getItem(key); } catch { return null; } }
function write(key, value) { try { localStorage.setItem(key, value); } catch { /* 只在本页有效 */ } }
function stamp() {
  const d = new Date(), pad = n => String(n).padStart(2, '0');
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
}
function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function setupTheme() {
  const buttons = document.querySelectorAll('[data-theme-choice]');
  function apply(theme) {
    document.documentElement.dataset.theme = theme;
    for (const button of buttons) button.setAttribute('aria-pressed', String(button.dataset.themeChoice === theme));
  }
  const savedTheme = read(THEME_KEY);
  apply(['arcade', 'neon'].includes(savedTheme) ? savedTheme : 'harbor');
  for (const button of buttons) {
    button.addEventListener('click', () => { apply(button.dataset.themeChoice); write(THEME_KEY, button.dataset.themeChoice); });
  }
}

let toastTimer;
function toast(message) {
  const box = $('toast');
  box.textContent = message;
  box.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => box.classList.remove('show'), 2400);
}
async function copyText(str) {
  try {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(str);
    } else {
      const tmp = document.createElement('textarea');
      tmp.value = str; tmp.style.position = 'fixed'; tmp.style.opacity = '0';
      document.body.append(tmp); tmp.select();
      if (!document.execCommand('copy')) throw Error('copy');
      tmp.remove();
    }
    toast('已复制到剪贴板');
  } catch { toast('复制失败：请手动选中并复制'); }
}
function download(name, mime, content) {
  const url = URL.createObjectURL(new Blob([content], {type:mime}));
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}
function csvCell(v) { return '"' + String(v ?? '').replaceAll('"', '""') + '"'; }
function fileName(name) { return String(name).replace(/[\\/:*?"<>|]+/g, '_').slice(0, 80); }

// ---------- 和程序通信 ----------

async function api(path, body) {
  const options = body === undefined ? {} : {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)};
  let response;
  try {
    response = await fetch(path, options);
  } catch {
    throw Error('连不上程序：请先启动工具，再用浏览器打开 /pcl');
  }
  let data = null;
  try { data = await response.json(); } catch { /* 下面统一报错 */ }
  if (!response.ok || !data || data.ok === false) throw Error(data?.error || '请求失败（HTTP ' + response.status + '）');
  return data;
}

let saving = 0;
function setSaveState(state, message) {
  const box = $('saveState');
  box.dataset.state = state;
  box.textContent = message || (state === 'saving' ? '保存中…' : '已保存到程序数据库');
}
// 写操作：显示保存状态；失败时提示并返回 null
async function save(path, body) {
  saving++;
  setSaveState('saving');
  try {
    const data = await api(path, body);
    if (--saving === 0) setSaveState('ok');
    if (target) target.updated_at = stamp();
    return data;
  } catch (err) {
    saving--;
    setSaveState('error', '没保存上：' + err.message);
    toast(err.message);
    return null;
  }
}
function requireTarget() {
  if (target) return true;
  toast('还没有连上程序，结果没法保存');
  return false;
}

// ---------- 每一项的结果 ----------

const EMPTY = Object.freeze({verdict:'', tester:'', tested_at:'', note:'', images:[], history:[]});
function result(id) { return results[id] || EMPTY; }
function ensureResult(id) { return results[id] ||= {verdict:'', tester:'', tested_at:'', note:'', images:[], history:[]}; }
function verdictOf(id) { return result(id).verdict; }

async function setVerdict(id, verdict) {
  if (!requireTarget()) return;
  const t = target, body = {target:t.id, item:id, verdict};
  const tester = $('testerName').value.trim();
  if (verdict && tester) body.tester = tester;
  const data = await save('/api/pcl/result', body);
  if (!data || target !== t) return;
  results[id] = data.result;
  updateTargetCounts();
  renderProgress(); renderList(); renderDetail(); renderChecklistCoverage(); renderFooter();
  toast(verdict ? id + ' 记为 ' + VERDICTS[verdict].label : id + ' 已撤销结论');
}

// 备注 / 测试人 / 时间：边打字边存，停 0.6 秒再发
const fieldTimers = new Map();
function queueField(id, field, value) {
  if (!target) return;
  const body = {target:target.id, item:id, [field]:value};
  const key = target.id + ':' + id + ':' + field;
  ensureResult(id)[field] = value;
  clearTimeout(fieldTimers.get(key)?.timer);
  const run = () => { fieldTimers.delete(key); return save('/api/pcl/result', body); };
  fieldTimers.set(key, {timer:setTimeout(run, 600), run, body});
}
function flushFields() {
  for (const {timer, run} of [...fieldTimers.values()]) { clearTimeout(timer); run(); }
}

// ---------- 渲染：进度、分类、列表 ----------

function renderProgress() {
  const items = selectedItems(), judged = items.filter(x => verdictOf(x.id)).length;
  const pct = items.length ? Math.round(judged / items.length * 100) : 0;
  $('pct').textContent = pct + '%';
  $('dial').style.setProperty('--p', pct + '%');
  $('progressCount').innerHTML = '已判定 <b>' + judged + '</b> / ' + items.length + ' 项' +
    (items.length < ITEMS.length ? '<br/>另有 ' + (ITEMS.length - items.length) + ' 项不测' : '');
  $('headerCount').textContent = judged + ' / ' + items.length;
}
function renderFooter() {
  $('lastSaved').textContent = target ? '「' + target.name + '」最后修改：' + target.updated_at : '尚无检查记录';
}
function renderCategories() {
  const nav = $('categoryList');
  nav.replaceChildren();
  for (const cat of CATS) {
    const count = cat.id === 'all' ? ITEMS.length : ITEMS.filter(x => x.tags.includes(cat.id)).length;
    const btn = el('button', 'filter-btn' + (activeCategory === cat.id ? ' active' : ''));
    btn.type = 'button';
    btn.setAttribute('aria-pressed', String(activeCategory === cat.id));
    const left = el('span', 'ftext'), mark = el('i', 'mark');
    mark.setAttribute('aria-hidden', 'true');
    left.append(mark, el('span', '', cat.name));
    btn.append(left, el('span', 'fcount', count));
    btn.addEventListener('click', () => { activeCategory = cat.id; renderCategories(); renderList(); });
    nav.append(btn);
  }
}

function matchesStatus(id) {
  if (activeStatus === 'all') return true;
  if (activeStatus === 'selected') return selection.has(id);
  if (activeStatus === 'pending') return selection.has(id) && !verdictOf(id);
  if (activeStatus === 'warn') return sourceWarnings.has(id);
  return verdictOf(id) === activeStatus;
}
function filtered() {
  return ITEMS.filter(x => (activeCategory === 'all' || x.tags.includes(activeCategory)) && matchesStatus(x.id) &&
    (!query || [x.title, x.jp, x.meaning, x.id, x.tags.join(' '), x.keys.join(' '), result(x.id).note].join(' ').toLowerCase().includes(query)));
}

// ---------- 选哪些项目测 ----------

function selectedItems() { return ITEMS.filter(x => selection.has(x.id)); }

function renderSelection() {
  const n = selection.size;
  $('selectCount').textContent = '已选 ' + n + ' / ' + ITEMS.length + ' 项';
  for (const id of ['runSelected', 'runSelectedTop']) {
    $(id).textContent = '▶ 测试选中的 ' + n + ' 项';
    $(id).disabled = !n;
  }
}

async function setSelection(ids) {
  if (!requireTarget()) return;
  selection = new Set(ids);
  target.selection = selectedItems().map(x => x.id);
  updateTargetCounts();
  renderSelection(); renderProgress(); renderList(); renderChecklistCoverage();
  await save('/api/pcl/selection', {target:target.id, items:target.selection});
}

// ---------- 测试：只跑选中的项目（或单独某一项）的自动规则 ----------

function hashText(text, hash = 2166136261) {
  for (let i = 0; i < text.length; i++) hash = Math.imul(hash ^ text.charCodeAt(i), 16777619);
  return hash >>> 0;
}
// 测试时代码 / 覆盖率是哪个版本：之后变了，就在结果上提示「需重测」
function computeSourceSig() {
  let hash = 2166136261;
  for (const file of sourceFiles) hash = hashText(file.name + '\n' + file.type + '\n' + file.code + '\n', hash);
  return sourceFiles.length + ':' + hash.toString(16);
}
function currentSig(id) {
  if (id !== 'T21') return sourceSig;
  return coverage ? 'cov:' + hashText(JSON.stringify([coverage.imported_at, coverage.source, coverage.totals])).toString(16) : '';
}
function isStale(id) { return !!runs[id] && runs[id].sig !== currentSig(id); }

function packFinding(finding) {
  return {status:finding.status, message:finding.message, extra:finding.extra || [],
    files:finding.files.map(({name, lines}) => ({name, count:lines.size, lines:[...lines].slice(0, 30)}))};
}

async function runTests(ids) {
  if (!requireTarget()) return;
  if (!ids.length) { toast('先勾选要测试的项目'); return; }
  const auto = ids.filter(id => AUTO_ITEMS.includes(id)), rules = auto.filter(id => id !== 'T21');
  const notes = [], found = new Map();
  if (rules.length && sourceFiles.length) {
    for (const finding of checkSources(sourceFiles)) {
      if (!rules.includes(finding.id)) continue;
      if (!found.has(finding.id)) found.set(finding.id, []);
      found.get(finding.id).push(packFinding(finding));
    }
  } else if (rules.length) {
    notes.push((ids.length === 1 ? '这一项' : rules.length + ' 项') + '要先在上方添加代码卡片');
  }
  if (auto.includes('T21')) {
    if (coverage) found.set('T21', [packFinding(coverageFinding(coverage))]);
    else notes.push('T21 要先在上方读取 clover.xml');
  }
  if (ids.length > auto.length) {
    notes.push((ids.length === 1 ? '这一项' : (ids.length - auto.length) + ' 项') + '没有自动规则，请按步骤人工确认后记结论');
  }
  const payload = {};
  for (const id of auto) {
    if (id === 'T21' ? !coverage : !sourceFiles.length) continue;
    const findings = found.get(id) || [];
    payload[id] = {status:Object.keys(SOURCE_STATUS).find(s => findings.some(f => f.status === s)) || 'none',
      sig:currentSig(id), findings};
  }
  const tested = Object.keys(payload);
  let summary = notes.join('；');
  if (tested.length) {
    const t = target;
    const data = await save('/api/pcl/runs', {target:t.id, tester:$('testerName').value.trim(), runs:payload});
    if (!data || target !== t) return;
    Object.assign(runs, data.runs);
    const counts = {};
    tested.forEach(id => { counts[runs[id].status] = (counts[runs[id].status] || 0) + 1; });
    summary = '测试了 ' + tested.length + ' 项：' + Object.keys(RUN_LABELS).filter(s => counts[s])
      .map(s => RUN_LABELS[s] + ' ' + counts[s]).join(' · ') + (summary ? '；' + summary : '');
    renderRuns();
  }
  $('runStatus').textContent = stamp().slice(11) + ' ' + summary;
  toast(summary);
  if (ids.length === 1) choose(ids[0]);
}

// 测试结果变了（测完 / 换对象 / 代码或覆盖率变了）：重画用到它的地方
function renderRuns() {
  const tested = Object.keys(runs), order = Object.keys(SOURCE_STATUS);
  sourceWarnings = new Set(tested.filter(id => runs[id].status === 'warn'));
  const findings = tested.flatMap(id => runs[id].findings.map(finding => ({...finding, id})))
    .sort((a, b) => order.indexOf(a.status) - order.indexOf(b.status) || a.id.localeCompare(b.id));
  $('sourceResults').replaceChildren(...findings.map(finding => findingCard(finding, true)));
  let summary = (sourceFiles.length ? '已添加 ' + sourceFiles.length + ' 个文件。' : '添加代码卡片后，') +
    '勾选要测试的项目再点「测试选中的项目」；也可以在某一项上点「▶ 测试」只测这一个。';
  if (tested.length) {
    const stale = tested.filter(isStale).length;
    summary = '已测试 ' + tested.length + ' 项：' + (sourceWarnings.size ? sourceWarnings.size + ' 项有疑点。' : '没有发现疑点。') +
      (findings.some(f => f.status === 'unknown') ? '有配置需要补充确认。' : '') +
      (stale ? stale + ' 项是按之前的代码 / 覆盖率测的，需重测。' : '') + (findings.length ? '下方结论仅适用于列出的代码位置。' : '');
  }
  $('sourceStatus').textContent = summary;
  renderSelection();
  renderList();
  renderAuto(ITEMS.find(x => x.id === selected) || ITEMS[0]);
  renderChecklistCoverage();
}
function sourcesChanged() {
  sourceSig = computeSourceSig();
  renderRuns();
}

function findingCount(finding) {
  return finding.files.reduce((n, file) => n + file.count, 0) || 1;
}
// 列表卡片上的一行测试结果摘要，按最严重的状态上色
function autoSummary(id) {
  const run = runs[id];
  if (!run) return AUTO_ITEMS.includes(id) ? {status:'none', text:'自动：还没测试'} : null;
  const counts = {};
  for (const finding of run.findings) counts[finding.status] = (counts[finding.status] || 0) + findingCount(finding);
  const parts = Object.keys(SOURCE_STATUS).filter(s => counts[s]).map(s => SOURCE_STATUS[s] + ' ' + counts[s]);
  return {status:run.status + (isStale(id) ? ' stale' : ''),
    text:'自动：' + (parts.join(' · ') || RUN_LABELS.none) + ' · ' + run.at + (isStale(id) ? ' · 之后有变化，需重测' : '')};
}

function taskReport(item) {
  const r = result(item.id), box = el('div', 'task-report');
  box.append(el('span', 'verdict-chip v-' + (r.verdict || 'none'), VERDICTS[r.verdict].label));
  const who = [r.tester, r.tested_at].filter(Boolean).join(' · ');
  if (who) box.append(el('span', 'report-who', who));
  const auto = autoSummary(item.id);
  if (auto) box.append(el('span', 'report-auto ' + auto.status, auto.text));
  if (r.images.length) box.append(el('span', 'report-shots', '📎 截图 ' + r.images.length));
  const note = r.note.trim().split('\n')[0];
  if (note) box.append(el('p', 'report-note', note));
  return box;
}

function renderList() {
  const subset = filtered(), root = $('taskList');
  root.replaceChildren();
  $('resultsCounter').textContent = '显示 ' + subset.length + ' / ' + ITEMS.length + ' 项';
  if (!subset.length) {
    const empty = el('div', 'empty');
    empty.innerHTML = '<h3>没有匹配的检查项目</h3><p>试试切换分类、测试结论或搜索关键词。</p>';
    root.append(empty);
    return;
  }
  for (const item of subset) {
    const verdict = verdictOf(item.id), on = selection.has(item.id);
    const card = el('article', 'task' + (verdict ? ' is-done v-' + verdict : '') + (on ? '' : ' is-off') +
      (selected === item.id ? ' selected' : ''));
    card.tabIndex = 0;
    card.setAttribute('role', 'button');
    card.setAttribute('aria-label', item.id + ' ' + item.title + '，' + (on ? '' : '不测，') + VERDICTS[verdict].label + '，查看详情');
    card.addEventListener('click', () => choose(item.id));
    card.addEventListener('keydown', e => {
      if (e.target === card && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); choose(item.id); }
    });
    const num = el('span', 'task-number', item.id.slice(1).padStart(2, '0'));
    const pick = el('button', 'check-control');
    pick.type = 'button';
    pick.setAttribute('role', 'checkbox');
    pick.setAttribute('aria-checked', String(on));
    pick.setAttribute('aria-label', '要测试 · ' + item.title);
    pick.title = on ? '已勾选：这一项要测。点一下改成不测' : '没勾选：这一项不测。点一下勾选';
    pick.innerHTML = '<svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3.3" stroke-linecap="round" stroke-linejoin="round"><path d="m5 12 4 4L19 6"/></svg>';
    pick.addEventListener('click', e => {
      e.stopPropagation();
      setSelection([...selection].filter(id => id !== item.id).concat(on ? [] : [item.id]));
    });
    const content = el('div', 'task-content');
    const sub = el('p', 'task-sub', item.meaning.length > 85 ? item.meaning.slice(0, 82) + '…' : item.meaning);
    const foot = el('div', 'task-footer'), tags = el('div', 'tagline');
    item.tags.forEach(tag => tags.append(el('span', 'tech-tag', LABELS[tag])));
    if (!on) tags.append(el('span', 'off-tag', '不测'));
    if (sourceWarnings.has(item.id)) tags.append(el('span', 'source-tag', item.id === 'T21' ? '覆盖率不足' : '源码有疑点'));
    foot.append(tags);
    if (AUTO_ITEMS.includes(item.id)) {
      const run = el('button', 'run-one', '▶ 测试');
      run.type = 'button';
      run.title = '只测试这一项';
      run.setAttribute('aria-label', '只测试这一项 · ' + item.title);
      run.addEventListener('click', e => { e.stopPropagation(); runTests([item.id]); });
      foot.append(run);
    } else {
      foot.append(el('span', 'task-state', '人工确认项'));
    }
    content.append(el('h3', 'task-title', item.title), sub, foot, taskReport(item));
    card.append(num, pick, content);
    root.append(card);
  }
}

// ---------- 渲染：右侧详情 + 测试结果报告 ----------

function choose(id) {
  if (selected === id) return;
  selected = id;
  renderList(); renderDetail();
  if (innerWidth <= 1080) $('detailPanel').scrollIntoView({behavior:'smooth', block:'start'});
}

function findingCard(finding, withLink) {
  const card = el('article', 'source-result ' + finding.status);
  card.append(el('span', 'source-result-status', SOURCE_STATUS[finding.status]));
  if (withLink) {
    const item = ITEMS.find(x => x.id === finding.id);
    const link = el('button', 'source-link', item.id + ' · ' + item.title);
    link.type = 'button';
    link.addEventListener('click', () => { choose(item.id); $('detailPanel').scrollIntoView({block:'nearest'}); });
    card.append(link);
  }
  card.append(el('p', '', finding.message));
  for (const {name, lines, count} of finding.files) {
    const list = el('ul');
    for (const [line, excerpt] of lines.slice(0, 5)) {
      const li = el('li');
      li.append('第 ' + line + ' 行 · ', el('code', '', excerpt));
      list.append(li);
    }
    if (count > 5) list.append(el('li', '', '另有 ' + (count - 5) + ' 行相关配置'));
    card.append(el('p', 'source-file-name', name), list);
  }
  if (finding.extra?.length) {
    const list = el('ul');
    finding.extra.forEach(text => list.append(el('li', '', text)));
    card.append(list);
  }
  return card;
}

function renderAuto(item) {
  const root = $('autoResults'), run = runs[item.id], auto = AUTO_ITEMS.includes(item.id), stale = isStale(item.id);
  root.replaceChildren();
  $('runOne').hidden = !auto;
  $('autoMeta').textContent = run ? '测试于 ' + run.at + (run.tester ? ' · ' + run.tester : '') +
    (stale ? ' · 之后' + (item.id === 'T21' ? '覆盖率' : '代码') + '有变化，需重测' : '') : '';
  $('autoMeta').classList.toggle('stale', stale);
  if (run?.findings.length) {
    run.findings.forEach(finding => root.append(findingCard(finding, false)));
    return;
  }
  root.append(el('p', 'auto-empty', !auto ? '这一项没有自动规则，请按上面的步骤人工确认后记结论。'
    : run ? '代码里没有找到和这一项相关的写法。'
      : item.id === 'T21' ? '还没测试。在上方「覆盖率」读取 PHPUnit 生成的 clover.xml，再点「测试这一项」。'
        : '还没测试。在上方添加代码卡片，再点「测试这一项」。'));
}

function renderShots(item) {
  const root = $('shotList');
  root.replaceChildren();
  for (const image of result(item.id).images) {
    const figure = el('figure', 'shot');
    const link = el('a');
    link.href = '/api/pcl/image?id=' + image.id;
    link.target = '_blank';
    link.rel = 'noopener';
    link.title = '在新标签页查看原图';
    const img = el('img');
    img.src = link.href;
    img.alt = image.name;
    img.loading = 'lazy';
    link.append(img);
    const remove = el('button', 'shot-del', '×');
    remove.type = 'button';
    remove.setAttribute('aria-label', '删除截图 ' + image.name);
    remove.addEventListener('click', async () => {
      if (!confirm('删除截图「' + image.name + '」？')) return;
      const t = target;
      if (!(await save('/api/pcl/image/delete', {id:image.id})) || target !== t) return;
      const list = result(item.id).images;
      list.splice(list.indexOf(image), 1);
      renderShots(item); renderList();
    });
    figure.append(link, el('figcaption', '', image.name), remove);
    root.append(figure);
  }
}

function renderHistory(item) {
  const root = $('historyList');
  root.replaceChildren();
  for (const change of [...result(item.id).history].reverse()) {
    root.append(el('li', '', change.at + ' · ' + (change.tester || '未填测试人') + ' · ' +
      (change.verdict ? '记为 ' + VERDICTS[change.verdict].label : '撤销结论')));
  }
}

function renderDetail() {
  const x = ITEMS.find(item => item.id === selected) || ITEMS[0], r = result(x.id);
  $('detailNo').textContent = 'CHECK / ' + x.id.slice(1).padStart(3, '0');
  $('detailTitle').textContent = x.title;
  $('detailDesc').textContent = x.tags.map(t => LABELS[t]).join('  ·  ');
  $('detailBadge').textContent = r.verdict ? '结论 · ' + VERDICTS[r.verdict].label : '待判定';
  $('detailBadge').className = 'detail-badge' + (r.verdict ? ' done v-' + r.verdict : '');
  $('original').textContent = x.jp;
  $('meaning').textContent = x.meaning;
  $('caution').textContent = x.note;
  const steps = $('steps');
  steps.replaceChildren();
  x.steps.forEach((s, i) => { const li = el('li'); li.append(el('span', 'step-num', i + 1), el('span', '', s)); steps.append(li); });
  const keywords = $('keywords');
  keywords.replaceChildren();
  x.keys.forEach(k => {
    const b = el('button', 'keyword', k);
    b.type = 'button'; b.title = '复制搜索关键词';
    b.addEventListener('click', () => copyText(k));
    keywords.append(b);
  });
  const snippet = $('snippetContainer');
  snippet.replaceChildren();
  if (x.code) {
    const wrap = el('div', 'codewrap'), copy = el('button', 'codecopy', '复制');
    copy.type = 'button';
    copy.addEventListener('click', () => copyText(x.code));
    wrap.append(el('div', 'code-label', '定位示例 / SEARCH REFERENCE'), copy, el('pre', '', x.code));
    snippet.append(wrap);
  }
  // 正在打字的输入框不覆盖，切换项目 / 对象时才重新填
  const shown = (target?.id ?? '') + ':' + x.id;
  for (const [id, value] of [['notes', r.note], ['reportTester', r.tester], ['reportTime', r.tested_at]]) {
    if (shownNoteFor !== shown || document.activeElement !== $(id)) $(id).value = value;
  }
  shownNoteFor = shown;
  document.querySelectorAll('.verdict-btn').forEach(b => b.setAttribute('aria-pressed', String(b.dataset.verdict === r.verdict)));
  renderAuto(x); renderShots(x); renderHistory(x);
  $('nextBtn').disabled = ITEMS.indexOf(x) === ITEMS.length - 1;
}

// ---------- 截图 ----------

function readDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}
async function addImages(files) {
  if (!requireTarget()) return;
  const t = target, item = selected;
  let added = 0;
  for (const file of files) {
    if (!/^image\/(?:png|jpeg|gif|webp)$/.test(file.type)) { toast('只能贴 PNG / JPEG / GIF / WebP 图片'); continue; }
    if (file.size > IMAGE_LIMIT) { toast('图片太大（最多 8 MB）'); continue; }
    const name = file.name && !/^image\.\w+$/.test(file.name) ? file.name : '截图 ' + stamp();
    const data = await save('/api/pcl/image', {target:t.id, item, name, data:await readDataUrl(file)});
    if (!data || target !== t) continue;
    ensureResult(item).images.push(data.image);
    added++;
  }
  if (!added) return;
  if (selected === item) renderShots(ITEMS.find(x => x.id === item));
  renderList();
  toast('已把 ' + added + ' 张截图加到 ' + item);
}

// ---------- 源码卡片 ----------

function persistSources() {
  if (!target) return Promise.resolve(null);
  return save('/api/pcl/sources/save', {target:target.id, sources:sourceFiles});
}

function renderSourceCards() {
  const root = $('sourceFiles');
  root.replaceChildren();
  $('sourceCardSection').hidden = sourceFiles.length === 0;
  $('reloadPaths').hidden = !sourceFiles.some(f => f.path);
  for (const file of sourceFiles) {
    const card = el('article', 'source-file'), header = el('div', 'source-card-header'), info = el('div');
    const meta = SOURCE_TYPES[file.type] + ' · ' + file.code.split('\n').length + ' 行';
    info.append(el('h4', 'source-card-title', file.name), el('p', 'source-card-meta', meta));
    if (file.path) info.append(el('p', 'source-card-meta source-card-path', file.path));
    const remove = el('button', 'btn', '删除');
    remove.type = 'button';
    remove.setAttribute('aria-label', '删除代码卡片 · ' + file.name);
    remove.addEventListener('click', () => {
      const index = sourceFiles.indexOf(file);
      sourceFiles.splice(index, 1);
      renderSourceCards(); sourcesChanged(); persistSources();
      const buttons = $('sourceFiles').querySelectorAll('.source-card-header .btn');
      (buttons[Math.min(index, buttons.length - 1)] || $('sourceCode')).focus();
    });
    header.append(info, remove);
    const preview = el('details', 'source-card-preview'), code = el('pre', '', file.code);
    code.tabIndex = 0;
    preview.append(el('summary', '', '查看代码'), code);
    card.append(header, preview);
    root.append(card);
  }
}

function updateSourceDraft() {
  const code = $('sourceCode').value, tooLong = code.length > SOURCE_LIMIT;
  $('addSource').disabled = tooLong || !code.trim();
  $('sourceDraftStatus').textContent = tooLong ? '超过 20 万字符，请分段添加。' : '';
}

// 按路径读进来的文件：同一路径已有卡片就换成最新内容
function mergeReadFiles(files) {
  let added = 0, updated = 0;
  for (const file of files) {
    const type = sourceType(file.code, file.name);
    const existing = sourceFiles.find(f => f.path && f.path.toLowerCase() === file.path.toLowerCase());
    if (existing) {
      if (existing.code !== file.code) updated++;
      Object.assign(existing, {code:file.code, type});
    } else {
      sourceFiles.push({name:file.name, path:file.path, code:file.code, type});
      added++;
    }
  }
  return {added, updated};
}
function readProblems(data) {
  const problems = [...data.missing, ...data.skipped];
  if (!problems.length) return '';
  return problems.length + ' 个没读：' + problems.slice(0, 3).map(p => p.path + '（' + p.reason + '）').join('；') +
    (problems.length > 3 ? '…' : '');
}
async function readPaths(paths, button) {
  if (!requireTarget()) return;
  button.disabled = true;
  $('pathStatus').textContent = '正在读取…';
  try {
    const data = await api('/api/pcl/sources/read', {paths});
    const {added, updated} = mergeReadFiles(data.files);
    renderSourceCards(); sourcesChanged();
    if (added || updated) await persistSources();
    const problems = readProblems(data);
    $('pathStatus').textContent = problems;
    toast('读取了 ' + data.files.length + ' 个文件：新增 ' + added + ' 个，内容有变化 ' + updated + ' 个' + (problems ? '；' + problems : ''));
    return true;
  } catch (err) {
    $('pathStatus').textContent = err.message;
    return false;
  } finally {
    button.disabled = false;
  }
}

// ---------- 覆盖率 ----------

function percentText(part) {
  const pct = coveragePercent(part);
  return pct === null ? '—' : pct + '%';
}
function stackedBar(counts, total) {
  const bar = el('div', 'cov-bar');
  bar.setAttribute('role', 'img');
  bar.setAttribute('aria-label', ['ok', 'ng', 'na'].map(v => VERDICTS[v].label + ' ' + counts[v]).join('，') + '，待判定 ' + counts['']);
  for (const v of ['ok', 'ng', 'na']) {
    if (!counts[v]) continue;
    const seg = el('span', 'seg-' + v);
    seg.style.width = counts[v] / total * 100 + '%';
    bar.append(seg);
  }
  return bar;
}
function verdictCounts(items) {
  const counts = {ok:0, ng:0, na:0, '':0};
  items.forEach(x => counts[verdictOf(x.id)]++);
  return counts;
}

// 只统计勾选要测的项目
function renderChecklistCoverage() {
  const root = $('checkCoverage');
  root.replaceChildren();
  const chosen = selectedItems(), counts = verdictCounts(chosen), judged = chosen.length - counts[''];
  if (!chosen.length) {
    root.append(el('p', 'focus-tip', '还没有勾选要测试的项目。在下面的检查项目里勾选后，这里统计它们的进度。'));
    return;
  }
  const stats = el('div', 'cov-stats');
  const stat = (label, value, className = '') => {
    const box = el('span', 'cov-stat ' + className);
    box.append(el('b', '', value), el('small', '', label));
    stats.append(box);
  };
  stat('已判定 ' + judged + ' / ' + chosen.length, Math.round(judged / chosen.length * 100) + '%');
  stat('OK', counts.ok, 'v-ok');
  stat('NG', counts.ng, 'v-ng');
  stat('不适用', counts.na, 'v-na');
  stat('待判定', counts['']);
  stat('通过率（OK / OK+NG）', counts.ok + counts.ng ? Math.round(counts.ok / (counts.ok + counts.ng) * 100) + '%' : '—');
  if (chosen.length < ITEMS.length) stat('没勾选，不测', ITEMS.length - chosen.length);
  root.append(stats, stackedBar(counts, chosen.length));
  const rows = el('div', 'cov-rows');
  for (const cat of CATS.slice(1)) {
    const items = chosen.filter(x => x.tags.includes(cat.id)), c = verdictCounts(items);
    const name = el('button', 'linkish', cat.name);
    name.type = 'button';
    name.title = '只看这一类';
    name.addEventListener('click', () => { activeCategory = cat.id; renderCategories(); renderList(); });
    rows.append(name, items.length ? stackedBar(c, items.length) : el('span', 'cov-none', '这一类没有勾选'),
      el('span', 'cov-num', (items.length - c['']) + ' / ' + items.length));
  }
  root.append(rows);
  const auto = el('div', 'cov-auto'), autoChosen = AUTO_ITEMS.filter(id => selection.has(id));
  const tested = autoChosen.filter(id => runs[id]), stale = tested.filter(isStale);
  auto.append(el('p', '', '勾选的 ' + chosen.length + ' 项里可自动测试 ' + autoChosen.length + ' 项，已测试 ' + tested.length + ' 项' +
    (stale.length ? '（其中 ' + stale.length + ' 项需重测）' : '') + '；其余 ' + (chosen.length - autoChosen.length) + ' 项需人工确认。'));
  const chips = el('div', 'auto-chips');
  for (const id of AUTO_ITEMS) {
    const status = runs[id]?.status || 'none';
    const chip = el('button', 'auto-chip ' + (runs[id] ? status : 'untested') + (selection.has(id) ? '' : ' off'), id);
    chip.type = 'button';
    chip.title = ITEMS.find(x => x.id === id).title + ' · ' + (runs[id] ? RUN_LABELS[status] : '还没测试') +
      (isStale(id) ? ' · 需重测' : '') + (selection.has(id) ? '' : ' · 没勾选');
    chip.addEventListener('click', () => choose(id));
    chips.append(chip);
  }
  auto.append(chips);
  root.append(auto);
}

function cardFileNames() {
  const names = new Set();
  for (const file of sourceFiles) {
    for (const value of [file.name, file.path]) {
      if (value) names.add(value.replace(/\\/g, '/').split('/').pop().toLowerCase());
    }
  }
  return names;
}

function renderCodeCoverage() {
  const root = $('codeCoverage');
  root.replaceChildren();
  $('coverageClear').hidden = !coverage;
  if (!coverage) {
    root.append(el('p', 'focus-tip', '还没有读取覆盖率。生成方法：vendor/bin/phpunit --coverage-clover build/clover.xml（需要 Xdebug 或 PCOV），然后填路径读取或选择文件。'));
    return;
  }
  const meters = el('div', 'cov-meters');
  for (const [key, label] of [['lines', '行'], ['methods', '方法'], ['conditionals', '条件分支']]) {
    const part = coverage.totals[key], pct = coveragePercent(part);
    const box = el('div', 'cov-meter' + (pct === 100 ? ' full' : pct !== null && pct < 100 ? ' low' : ''));
    box.append(el('strong', '', percentText(part)), el('span', '', label + ' ' + (part.total ? part.covered + ' / ' + part.total : '无数据')));
    meters.append(box);
  }
  root.append(meters);
  root.append(el('p', 'cov-meta', '来源：' + coverage.source + (coverage.generated ? ' · 生成于 ' + coverage.generated : '') +
    ' · 读取于 ' + coverage.imported_at + (coverage.prefix ? ' · 路径前缀 ' + coverage.prefix : '')));
  const names = cardFileNames();
  const related = coverage.files.filter(f => names.has(f.name.replace(/\\/g, '/').split('/').pop().toLowerCase()));
  if (related.length) {
    const label = el('label', 'cov-filter'), box = el('input');
    box.type = 'checkbox';
    box.checked = coverageOnlyCards;
    box.addEventListener('change', () => { coverageOnlyCards = box.checked; renderCodeCoverage(); });
    label.append(box, ' 只看代码卡片里的文件（' + related.length + ' 个）');
    root.append(label);
  }
  const files = (coverageOnlyCards && related.length ? related : coverage.files).slice()
    .sort((a, b) => (coveragePercent(a.lines) ?? 101) - (coveragePercent(b.lines) ?? 101) || a.short.localeCompare(b.short));
  const wrap = el('div', 'cov-table-wrap'), table = el('table', 'cov-table');
  table.innerHTML = '<thead><tr><th scope="col">文件</th><th scope="col">行</th><th scope="col">方法</th><th scope="col">没执行到的行 / 方法</th></tr></thead>';
  const body = el('tbody');
  for (const file of files.slice(0, 300)) {
    const full = file.lines.covered === file.lines.total && file.conditionals.covered === file.conditionals.total;
    const row = el('tr', full ? 'full' : '');
    const name = el('td', 'file', file.short);
    name.title = file.name;
    const miss = [];
    if (file.uncovered.length) miss.push('第 ' + coverageRanges(file.uncovered, 12) + ' 行');
    if (file.partial?.length) miss.push('分支没走全：第 ' + coverageRanges(file.partial, 6) + ' 行');
    if (file.uncovered_methods.length) miss.push('方法 ' + file.uncovered_methods.slice(0, 6).join(', ') + (file.uncovered_methods.length > 6 ? '…' : ''));
    row.append(name, el('td', 'num', percentText(file.lines)), el('td', 'num', percentText(file.methods)), el('td', 'miss', miss.join('；') || '✓'));
    body.append(row);
  }
  table.append(body);
  wrap.append(table);
  root.append(wrap);
  if (files.length > 300) root.append(el('p', 'cov-meta', '只列出覆盖率最低的 300 个文件，共 ' + files.length + ' 个。'));
}

async function importCoverage(body) {
  if (!requireTarget()) return;
  const t = target;
  const data = await save('/api/pcl/coverage', {target:t.id, ...body});
  if (!data || target !== t) return;
  coverage = data.coverage;
  renderCodeCoverage(); renderRuns();
  toast('已读取覆盖率：行 ' + percentText(coverage.totals.lines) + '，' + coverage.files.length + ' 个文件。勾选 T21 后点测试，或在 T21 上点「▶ 测试」');
}

// ---------- 检查对象 ----------

function updateTargetCounts() {
  if (!target) return;
  target.counts = {ok:0, ng:0, na:0};
  selectedItems().forEach(x => { if (verdictOf(x.id)) target.counts[verdictOf(x.id)]++; });
  renderTargetSelect();
}
function renderTargetSelect() {
  const select = $('targetSelect');
  select.replaceChildren(...targets.map(t => {
    const judged = t.counts.ok + t.counts.ng + t.counts.na, total = t.selection ? t.selection.length : ITEMS.length;
    const option = el('option', '', t.name + '（已判定 ' + judged + ' / ' + total + (t.counts.ng ? ' · NG ' + t.counts.ng : '') + '）');
    option.value = t.id;
    option.selected = t.id === target?.id;
    return option;
  }));
}

async function openTarget(id) {
  flushFields();
  const data = await api('/api/pcl/target?id=' + id);
  target = data.target;
  targets = targets.map(t => t.id === target.id ? target : t);
  results = data.results;
  coverage = data.coverage;
  sourceFiles = data.sources;
  runs = data.runs;
  selection = new Set(target.selection || ITEMS.map(x => x.id));
  nextSourceNumber = sourceFiles.length + 1;
  coverageOnlyCards = false;
  shownNoteFor = '';
  $('runStatus').textContent = '';
  write(TARGET_KEY, String(target.id));
  renderTargetSelect();
  renderSourceCards();
  sourcesChanged();
  renderAll();
  $('detailBody').scrollTop = 0;
  setSaveState('ok');
}

async function loadTargets() {
  targets = (await api('/api/pcl/targets')).targets;
  if (!targets.length) targets = [(await api('/api/pcl/targets/create', {name:'未命名检查对象'})).target];
  const remembered = Number(read(TARGET_KEY));
  await openTarget((targets.find(t => t.id === remembered) || targets[0]).id);
}

// 新建 / 改名共用页面上的一行输入框（有的内嵌浏览器不支持 prompt()）
let targetFormMode = '';
function showTargetForm(mode) {
  if (mode === 'rename' && !requireTarget()) return;
  targetFormMode = mode;
  $('targetFormLabel').textContent = mode === 'rename' ? '检查对象的新名称' : '新检查对象名称';
  $('targetNameInput').value = mode === 'rename' ? target.name : '';
  $('targetForm').hidden = false;
  $('targetNameInput').focus();
  $('targetNameInput').select();
}
function hideTargetForm() {
  $('targetForm').hidden = true;
  targetFormMode = '';
}
async function submitTargetForm() {
  const name = $('targetNameInput').value.trim();
  if (!name) return;
  if (targetFormMode === 'rename') {
    if (!requireTarget()) return;
    if (name !== target.name) {
      const data = await save('/api/pcl/targets/update', {target:target.id, name});
      if (!data) return;
      Object.assign(target, data.target);
      renderTargetSelect(); renderFooter();
    }
    hideTargetForm();
    return;
  }
  const data = await save('/api/pcl/targets/create', {name});
  if (!data) return;
  hideTargetForm();
  targets.unshift(data.target);
  await openTarget(data.target.id).catch(err => toast(err.message));
  toast('已新建「' + data.target.name + '」');
}
async function deleteTarget() {
  if (!requireTarget()) return;
  if (!confirm('删除「' + target.name + '」？\n22 项结论、备注、截图、源码卡片和覆盖率都会删掉，不能撤销。建议先导出 JSON。')) return;
  const removed = target;
  if (!(await save('/api/pcl/targets/delete', {target:removed.id}))) return;
  target = null;
  targets = targets.filter(t => t.id !== removed.id);
  if (!targets.length) {
    const data = await save('/api/pcl/targets/create', {name:'未命名检查对象'});
    if (data) targets = [data.target];
  }
  if (targets.length) await openTarget(targets[0].id).catch(err => toast(err.message));
  toast('已删除「' + removed.name + '」');
}

// ---------- 导出 / 导入 ----------

function autoText(id) {
  const run = runs[id];
  if (!run) return AUTO_ITEMS.includes(id) ? '还没测试' : '人工确认项';
  return [RUN_LABELS[run.status] + '（' + run.at + (isStale(id) ? '，之后有变化需重测' : '') + '）',
    ...run.findings.map(f => SOURCE_STATUS[f.status] + '：' + f.message)].join('\n');
}
async function exportJson() {
  if (!requireTarget()) return;
  flushFields();
  try {
    const {data} = await api('/api/pcl/export?id=' + target.id);
    const byId = new Map(data.items.map(entry => [entry.id, entry]));
    data.items = ITEMS.map(x => ({
      ...(byId.get(x.id) || {id:x.id, done:false, verdict:'', tester:'', testedAt:'', note:'', history:[], images:[]}),
      title:x.title, japanese:x.jp, selected:selection.has(x.id), autoCheck:autoText(x.id)
    }));
    download('PCL_单体测试报告_' + fileName(target.name) + '_' + new Date().toISOString().slice(0, 10) + '.json',
      'application/json;charset=utf-8', JSON.stringify(data, null, 2));
    toast('已导出 JSON 报告');
  } catch (err) { toast('导出失败：' + err.message); }
}
function exportCsv() {
  const header = ['编号', '分类', '日文原文', '中文检查项', '是否测试', '结论', '测试人', '测试时间', '自动测试结果', '截图数', '备注 / 证据说明'];
  const lines = [header.map(csvCell).join(',')];
  for (const x of ITEMS) {
    const r = result(x.id);
    lines.push([x.id, x.tags.map(t => LABELS[t]).join('/'), x.jp, x.title, selection.has(x.id) ? '测' : '不测',
      VERDICTS[r.verdict].label, r.tester, r.tested_at, autoText(x.id), r.images.length, r.note].map(csvCell).join(','));
  }
  download('PCL_单体测试清单_' + fileName(target?.name || '') + '_' + new Date().toISOString().slice(0, 10) + '.csv',
    'text/csv;charset=utf-8', '\ufeff' + lines.join('\r\n'));
  toast('已导出可用 Excel 打开的 CSV');
}
async function importJson(data, message) {
  if (!requireTarget()) return false;
  if (data?.format !== 'pcl-qa-checklist' || ![1, 2].includes(data.version) || !Array.isArray(data.items)) {
    toast('导入失败：文件格式不正确');
    return false;
  }
  if (!confirm(message)) return false;
  flushFields();
  const t = target;
  const done = await save('/api/pcl/import', {target:t.id, data});
  if (!done) return false;
  await openTarget(t.id).catch(err => toast(err.message));
  toast('导入成功：' + done.items + ' 项' + (done.images ? '，截图 ' + done.images + ' 张' : ''));
  return true;
}

// 旧版页面把进度存在这个浏览器里：第一次打开新版时问一次要不要导入
async function offerLegacyImport() {
  if (read(LEGACY_ASKED_KEY) || !target) return;
  let old = null;
  try { old = JSON.parse(read(LEGACY_KEY) || 'null'); } catch { /* 坏数据不导入 */ }
  const done = old?.done && typeof old.done === 'object' ? old.done : {};
  const notes = old?.notes && typeof old.notes === 'object' ? old.notes : {};
  const doneCount = ITEMS.filter(x => done[x.id] === true).length;
  const noteCount = ITEMS.filter(x => typeof notes[x.id] === 'string' && notes[x.id].trim()).length;
  if (!doneCount && !noteCount) return;
  write(LEGACY_ASKED_KEY, '1');
  const data = {format:'pcl-qa-checklist', version:1,
    items:ITEMS.map(x => ({id:x.id, done:done[x.id] === true, note:typeof notes[x.id] === 'string' ? notes[x.id] : ''}))};
  await importJson(data, '发现这个浏览器里旧版检查台的记录（已完成 ' + doneCount + ' 项、备注 ' + noteCount + ' 条）。\n' +
    '导入到当前检查对象「' + target.name + '」吗？已完成会记为 OK，当前对象已有的结论会被覆盖。\n（取消后不再询问，旧记录仍留在浏览器里）');
}

// ---------- 事件 ----------

function renderAll() {
  renderProgress(); renderCategories(); renderList(); renderDetail(); renderFooter();
  renderChecklistCoverage(); renderCodeCoverage(); renderSelection();
}

function bindEvents() {
  $('targetSelect').addEventListener('change', e => openTarget(Number(e.target.value)).catch(err => toast(err.message)));
  $('targetNew').addEventListener('click', () => showTargetForm('create'));
  $('targetRename').addEventListener('click', () => showTargetForm('rename'));
  $('targetForm').addEventListener('submit', e => { e.preventDefault(); submitTargetForm(); });
  $('targetFormCancel').addEventListener('click', hideTargetForm);
  $('targetNameInput').addEventListener('keydown', e => { if (e.key === 'Escape') hideTargetForm(); });
  $('targetDelete').addEventListener('click', deleteTarget);
  $('testerName').addEventListener('input', e => write(TESTER_KEY, e.target.value.trim()));

  $('sourceCode').addEventListener('input', updateSourceDraft);
  $('sourceForm').addEventListener('submit', event => {
    event.preventDefault();
    const code = $('sourceCode').value;
    if (!code.trim() || code.length > SOURCE_LIMIT) return;
    const filename = $('sourceName').value.trim(), requested = $('sourceType').value;
    sourceFiles.push({code, name:filename || '文件 ' + nextSourceNumber, path:'',
      type:requested === 'auto' ? sourceType(code, filename) : requested});
    nextSourceNumber++;
    $('sourceCode').value = '';
    $('sourceName').value = '';
    $('sourceType').value = 'auto';
    updateSourceDraft();
    renderSourceCards(); sourcesChanged(); persistSources();
    $('sourceCode').focus();
  });
  $('pathForm').addEventListener('submit', async event => {
    event.preventDefault();
    const paths = $('sourcePaths').value;
    if (!paths.trim()) { $('pathStatus').textContent = '请填写路径，每行一个'; return; }
    if (await readPaths(paths, $('readPaths'))) $('sourcePaths').value = '';
  });
  $('reloadPaths').addEventListener('click', () =>
    readPaths(sourceFiles.filter(f => f.path).map(f => f.path).join('\n'), $('reloadPaths')));
  $('clearSources').addEventListener('click', () => {
    if (!confirm('清空全部 ' + sourceFiles.length + ' 张代码卡片？')) return;
    sourceFiles = [];
    renderSourceCards(); sourcesChanged(); persistSources();
  });

  $('coverageForm').addEventListener('submit', event => {
    event.preventDefault();
    const path = $('coveragePath').value.trim();
    if (!path) { toast('请填写 clover.xml 的路径，或点「选择文件…」'); return; }
    importCoverage({path});
  });
  $('coveragePick').addEventListener('click', () => $('coverageFile').click());
  $('coverageFile').addEventListener('change', async e => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (file) importCoverage({xml:await file.text(), name:file.name});
  });
  $('coverageClear').addEventListener('click', async () => {
    if (!requireTarget() || !confirm('清除这个检查对象的代码覆盖率？')) return;
    const t = target;
    if (!(await save('/api/pcl/coverage/clear', {target:t.id})) || target !== t) return;
    coverage = null;
    renderCodeCoverage(); renderRuns();
  });

  // 全选 / 全不选只动当前显示的项目（配合左侧分类、搜索用）
  $('selectAll').addEventListener('click', () => setSelection([...selection, ...filtered().map(x => x.id)]));
  $('selectNone').addEventListener('click', () => {
    const shown = new Set(filtered().map(x => x.id));
    setSelection([...selection].filter(id => !shown.has(id)));
  });
  $('selectAuto').addEventListener('click', () => setSelection(AUTO_ITEMS));
  for (const id of ['runSelected', 'runSelectedTop']) {
    $(id).addEventListener('click', () => runTests(selectedItems().map(x => x.id)));
  }
  $('runOne').addEventListener('click', () => runTests([selected]));

  document.querySelectorAll('.verdict-btn').forEach(button => button.addEventListener('click', () =>
    setVerdict(selected, verdictOf(selected) === button.dataset.verdict ? '' : button.dataset.verdict)));
  $('nextBtn').addEventListener('click', () => {
    const index = ITEMS.findIndex(x => x.id === selected);
    if (index < ITEMS.length - 1) { selected = ITEMS[index + 1].id; renderList(); renderDetail(); $('detailBody').scrollTop = 0; }
  });
  $('notes').addEventListener('input', e => { queueField(selected, 'note', e.target.value); renderList(); });
  $('reportTester').addEventListener('input', e => { queueField(selected, 'tester', e.target.value.trim()); renderList(); });
  $('reportTime').addEventListener('input', e => { queueField(selected, 'tested_at', e.target.value.trim()); renderList(); });

  $('shotPick').addEventListener('click', () => $('shotFile').click());
  $('shotFile').addEventListener('change', e => { const files = [...e.target.files]; e.target.value = ''; addImages(files); });
  const drop = $('shotDrop');
  drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('dragging'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('dragging'));
  drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('dragging'); addImages([...e.dataTransfer.files]); });
  // 在页面任意处粘贴截图 → 加到当前选中的检查项（粘贴代码的输入框除外）
  document.addEventListener('paste', e => {
    if (['sourceCode', 'sourcePaths', 'sourceName', 'coveragePath'].includes(document.activeElement?.id)) return;
    const files = [...(e.clipboardData?.files || [])].filter(f => f.type.startsWith('image/'));
    if (!files.length) return;
    e.preventDefault();
    addImages(files);
  });

  $('searchInput').addEventListener('input', e => { query = e.target.value.trim().toLowerCase(); renderList(); });
  document.querySelectorAll('[data-status]').forEach(btn => btn.addEventListener('click', () => {
    activeStatus = btn.dataset.status;
    document.querySelectorAll('[data-status]').forEach(b => {
      const on = b === btn;
      b.classList.toggle('active', on);
      b.setAttribute('aria-pressed', String(on));
    });
    renderList();
  }));
  $('exportBtn').addEventListener('click', exportJson);
  $('csvBtn').addEventListener('click', exportCsv);
  $('importBtn').addEventListener('click', () => $('importFile').click());
  $('importFile').addEventListener('change', async e => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;
    let data;
    try { data = JSON.parse(await file.text()); } catch { toast('导入失败：不是有效的 JSON'); return; }
    importJson(data, '导入将覆盖「' + target?.name + '」的 22 项结论、备注、截图、历史和测试结果' +
      (data?.version === 2 ? '（文件里带的勾选、源码卡片和覆盖率也会换掉）' : '') + '，继续吗？');
  });
  $('resetBtn').addEventListener('click', async () => {
    if (!requireTarget()) return;
    if (!confirm('清空「' + target.name + '」22 项的结论、备注、截图、历史和测试结果？勾选、源码卡片和覆盖率保留。建议先导出备份。')) return;
    const t = target;
    if (!(await save('/api/pcl/reset', {target:t.id}))) return;
    await openTarget(t.id).catch(err => toast(err.message));
    toast('已清空测试结果');
  });
  // 关页面前把还没发出去的备注发出去
  addEventListener('pagehide', () => {
    for (const {timer, body} of fieldTimers.values()) {
      clearTimeout(timer);
      navigator.sendBeacon?.('/api/pcl/result', new Blob([JSON.stringify(body)], {type:'text/plain'}));
    }
    fieldTimers.clear();
  });
}

setupTheme();
bindEvents();
$('testerName').value = read(TESTER_KEY) || '';
updateSourceDraft();
renderAll();
loadTargets().then(offerLegacyImport).catch(err => {
  setSaveState('error', err.message);
  $('storageState').textContent = '没有连上程序，结果不会保存。';
  toast(err.message);
});
})();
