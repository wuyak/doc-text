# doc-text 开发方案

更新：2026-09-15。本文统一记录 doc-text 的目标、已确认行为、实现要求和验收依据。项目已从 legacy-doc 二开 worktree 迁为独立子目录和 Git 仓库；保留纯 Python、零运行时依赖和最小文字提取范围。

## 1. 目标、范围与公开行为

在 `legacy-doc` 内补齐规范文字定位、正文选择、段落及表格识别、文字拼接，使提取结果遵循已选定的 DOCX 文本规则。保持纯 Python、零运行时依赖，复用现有 OLE 读取和元数据能力。

本方案只涉及独立文字提取库。平台已有 DOCX 函数作为固定的输出参考；平台接入、提示链路、界面文案和提示类型留待后续另行设计，不属于本次实现或验收范围。DOCX 参考源码不成为库的运行时依赖。

以下决定已确认：

- **统一替换底层。** 所有入口使用 FIB 指定的文字记录，不保留旧 CLX 扫描路径。
- **默认直接使用新输出。** 现有 `extract_text()` 调用直接返回本方案定义的结果，不增加切回旧文字压缩行为的开关；不引入此前讨论的多模式选项。
- **沿用结果类型。** 继续返回 `DocExtractionResult`，使用现有 `text`、`metadata`、`warnings` 等字段，不新增公开的 Word 文档或表格对象模型。
- **合法空正文返回空字符串。** 按本方案选择和拼接后没有文字时，返回 `text=""`，不将它当作文件损坏，也不把提示文案填入正文。
- **影响所选文字完整性的损坏使整体失败。** 不将残缺文本作为成功结果；可选元数据损坏沿用 warning。

这次改变会影响旧调用方实际得到的文字：恢复表格分隔、保留主体空段落和连续空格、排除未选文档部分，并停止按 PAGE 字样删除所选正文内容。发布说明应明确记录这些变化；调用函数和结果载体不因此另建一套。

## 2. 固定源码与已有差距

| 用途 | 基线 | 本地位置 |
| --- | --- | --- |
| 原始基线 | `legacy-doc` 0.2.1；`33aac7cace0dc30d95e671df5aaa825e4c2223a1` | `/Users/gearup/workspaces/legacy-doc-development/legacy-doc` |
| 二开 worktree | 分支 `codex/docx-aligned-text`，从上述基线创建 | `/Users/gearup/workspaces/legacy-doc-development/legacy-doc-docx-aligned` |
| DOCX 输出参考 | 平台提交 `8c7d209bc03bf55103b1428f177e75f663185af2` | `/Users/gearup/workspaces/ai交互平台/Interactive-platform/backend/open_webui/retrieval/loaders/text_only.py` |
| DOCX 结构源码参考 | `python-docx` v1.2.0；`e45454602b53e8e572b179ccf1c91093ec9f4ed7` | `/Users/gearup/workspaces/legacy-doc-development/python-docx` |

输出参考是自写 XML 提取器的 `_clean_text()`、`_wml_inline_text()`、`_docx_table_text()`、`_docx_to_documents()`。它只读取 `word/document.xml`，按节点顺序取文字，不执行 Word 排版或判断最终显示状态。`python-docx` 用来理解段落和表格组织，不能替代这份输出基准：它的合并格展开、字段和修订文字读取均有不同。

`legacy-doc` 官方仓库为 [lcorrigan704/legacy-doc](https://github.com/lcorrigan704/legacy-doc)，MIT，Python >=3.11；本地参考库来自 [python-openxml/python-docx](https://github.com/python-openxml/python-docx)，同为 MIT。二进制格式以微软 [MS-DOC](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/ccd7b486-7881-484c-a137-51170af7cc22) 为依据，形状文字的相关记录参考 MS-ODRAW。

| 基线源码 | 二开前行为 | 本次改动 |
| --- | --- | --- |
| `word.py` 的 `_find_clx()` | 扫描 Table stream，取首个形似 piece table 的片段 | 按 FIB 的指定位置读取 CLX |
| `_extract_text_from_piece_table()` | 解码并拼接所有片段，不保存 CP、FC、PRM；部分越界片段静默跳过 | 保存位置和属性引用，按选择范围读取，必需记录损坏时报错 |
| `word.py` 的范围判断 | 将 0x18、0x1C 保留字段用作文字范围 | 使用 PCD、cbMac 和各 stream 边界 |
| 正文与表格处理 | 没有正文选择或表格属性解析；只有控制符替换 | 新增 story 选择、文本框关联、段落及表格关系 |
| `normalize.py` | 合并空格/TAB、丢空行、按正则删除 PAGE、过滤私用字符 | 按结构拼接并执行确定的空白规则 |
| 空正文处理 | 规范化后为空则抛异常 | 合法空正文返回 `text=""` |
| `tests/fixtures.py` | 简化 FIB，按 Python `len()` 计算 CP | 补规范字段，并按 UTF-16 单元计数 |

## 3. 二进制解析与输出设计

### 3.1 按 FIB 定位文字，保存字符与字节位置

| 规则 | 对实现的要求 | 官方依据 |
| --- | --- | --- |
| FIB 记录版本及后续结构位置 | 校验签名、计数和有效版本；考虑 nFibNew 的覆盖；按支持版本解释结构 | [Fib](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9aeaa2e7-4a45-468e-ab13-3f6193eb9394)、[FibBase](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/26fb6c06-4e5c-4778-ab4e-edbf26a545bb) |
| fcClx/lcbClx 指向 CLX | 在对应 0Table/1Table stream 中按指定偏移和长度读取，解析 RgPrc 与 Pcdt/PlcPcd | [取文本算法](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/01d5d8c4-cf9c-4ef9-80fd-439e763cfe01)、[Clx](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/bad26767-b575-44d3-9da3-96378d56ce14) |
| CP 是文档字符位置，FC 是存储位置 | 保留两者映射，按 CP 顺序处理，不能按物理地址排序文字 | [取文本算法](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/01d5d8c4-cf9c-4ef9-80fd-439e763cfe01) |
| 压缩/非压缩 piece 的位置和宽度不同 | 压缩位置值先除以 2，每 CP 一字节；非压缩每 CP 两字节；内部 FC 统一为实际字节偏移 | [FcCompressed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/aa2e55a2-f4f2-4795-bab5-6d9d7a0ed249) |
| cbMac 限定有效 WordDocument 内容 | 对该 stream 的文字和结构引用同时检查 cbMac 与实际长度；Table/Data stream 各自按自己的长度检查 | [FibRgLw97](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/37713d3c-a0c8-40f5-821f-bc9622c7de48) |

基线代码用作 fc_min/fc_mac 的 FIB 0x18、0x1C，在所依据规范中是未定义保留字段。新实现不能继续用它们判断哪些文字有效。Word 6/95 不在本次新增支持范围内；支持版本的清单和各版本长度条件要写进源码测试，不能只使用“版本大于某值”这样的宽泛判断。

建议使用小型私有记录保存信息：

- `FibInfo`：有效版本、所用 Table stream、cbMac、各文档部分的长度及定位阶段需要的结构位置。
- `Piece`：`cp_start`、`cp_end`、实际字节偏移 `fc`、压缩标志和 `prm`。保留 PRM 供后续段落属性解析使用。
- CLX 属性组：保存 Prm1 后续引用所需的 Prc 内容及索引关系。
- 区间读取能力：输入合法 CP 半开区间 `[start, end)`，与各 piece 求交后定位字节；选择哪个区间由调用方决定。

字段名为实现建议，不构成新的公开 API 承诺。

文字读取应遵守：

1. 非压缩文字在定位阶段按 UTF-16 的 16 位单元计数，不能拿 Python 字符串长度当作 CP 数。
2. 先计算区间和结构位置，再解码；不能先清洗字符串后据其长度回推原文件位置。
3. 保留跨 piece 的有效代理对，正确读取 emoji 等非 BMP 字符；不能分别使用替换解码后再拼接，造成有效字符丢失。
4. 压缩文字按 MS-DOC 指定映射解释。对于位于文本范围内但不显示的控制符，定位阶段保留其值，交给后续结构和清洗步骤处理。
5. 不在定位阶段合并空格、删除 PAGE 字样、解释字段结果或改变段落分隔。

### 3.2 选择正文，并关联另存的正文文字

| 内容 | DOCX 当前行为 | DOC 需要对应的行为 |
| --- | --- | --- |
| 主体段落、表格 | 读取主体中的段落与表格文字 | 从主文档字符范围读取，并恢复段落、表格关系 |
| 独立页眉、页脚 | 不打开对应 XML | 不提取 header story；页眉中的文本框也不纳入 |
| 独立脚注、尾注、批注 | 不打开对应 XML | 不提取对应 story 的说明文字；正文中的普通字面内容不因此删除 |
| 正文文本框 | 递归可达的文字节点会被读取 | 按正文形状锚点找回文本框内容，不能仅截取 `ccpText` 后就宣称完成 |
| 字段代码及结果 | `instrText` 与 `t` 都读取 | 保留已存储的代码和结果文字；不计算字段，不只取结果，不正则删除 PAGE |
| 插入、删除修订文字 | `t` 与 `delText` 都读取，不检查修订状态 | 保留目标正文范围内已存储的相应文字，不主动接受或拒绝修订 |
| 隐藏文字 | 不检查文字的隐藏属性 | 不因 `CFVanish`、`CFFldVanish` 等显示属性过滤已选文字 |
| 图片、外部对象 | 不做 OCR，不跟随外部关系提取对象 | 不因图片或对象存在而增加 OCR、联网或附件解包流程 |
| 艺术字 | 可达的 `t` 节点会读取；仅在属性中的字符串不会读取 | 必须核对具体存储形式，不能把所有艺术字概括为“全取”或“全不取” |

这里的“正文”是提取范围，不能简单等同于 DOC 规范中的 main document story。后者是 DOC 的一个连续字符区间，正文文本框通常另存。微软明确规定主文档范围为 `[0, ccpText)`，并包含指向其他内容的锚点。[Main Document](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/f426d9a2-004d-418e-8d8c-e7fd88e7c48e)、[Textboxes](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/f87b3560-2c23-4d10-9751-ff141d307308)。

1. **主体范围。** 从 FIB 读取 `ccpText` 等各 story 长度，按规范确定字符区间；正文、文本框的 CP 不能混用各自的局部坐标。所有加法和引用在读取前检查边界。[FibRgLw97](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/37713d3c-a0c8-40f5-821f-bc9622c7de48)。
2. **正文形状与文字。** 读取正文形状锚点，关联 OfficeArt 形状标识与文本框记录，再读取对应文本框区间。按锚点的正文位置加入内容，不能把所有文本框统一追加到文末。未被正文引用的闲置文本框记录不得顺带输出。[PlcfSpa](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/e98fe93c-d57f-4d9b-a519-30e6acbdd414)、[PlcftxbxTxt](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9f17f89a-64d6-4546-a811-e62e1d238102)、[FTXBXS](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/d0a81fe8-f9e8-4130-9233-84732e3d9420)。
3. **普通文字与控制记录。** 字段边界和对象锚点不能作为可见字符直接输出；也不能删除它们附近的普通文字。必要时使用字符属性 `CFSpec` 识别特殊字符。读取结构标记与过滤字段代码是两回事。[字符属性读取](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/be58bf9c-d1d3-40cc-91ee-36452d7939b2)、[Fields](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/751b09bb-72f0-45ef-8e87-666dea68219f)。

实现时需把关联落实为可检查的字段：FIB 的 `fcPlcfSpaMom/lcbPlcfSpaMom` 定位正文锚点；`fcPlcftxbxTxt/lcbPlcftxbxTxt` 定位文本框范围；使用 `SPA.lid`、OfficeArt 的 `spid`、`FTXBXS.lid` 关联形状；核对 `lTxid` 高 16 位所指的 FTXBXS 一基索引。`fReusable` 槽及最后的特殊记录不能当作实际文本框输出。只解析所需标识和引用，不计算几何位置及绕排。[lTxid](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/6f50606e-9a41-4555-961f-95317c8ed15c)。

字段指令、结果以及标为删除修订的文字仍存于所属 story，通常不需要另找文字内容。字段位置表描述边界，不代替文字片段；输出消费 `0x13/0x14/0x15` 等结构标记，保留其间已存储的文字。`CFSpec` 表示特殊含义，不是隐藏开关。[Character Properties](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/7022285b-9621-42e9-ad4d-4e02c115ef18)。

### 3.3 读取段落属性，恢复行、格和嵌套关系

DOC 的第一层表格用特殊段落标记区分格与行；嵌套表格还要结合层级属性。段落可能跨 piece，应找到真实结尾，再读取适用属性。仅把 `0x07` 换成 TAB，无法判断哪一处应结束外层单元格或表格行。[Tables](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/5b45f0e7-7760-4fdb-af88-0146de2feb4c)、[段落边界](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/30461a5b-e3ad-44cd-a3fe-038f86639b13)。

1. 从 FIB 的 `fcPlcfBtePapx/lcbPlcfBtePapx` 定位段落属性索引，再定位 512 字节 PAPX FKP 页。遇到正文 `0x0C` 时，结合 `PlcfSed.aCP` 区分分节结束与段内手动分页；后者只在段内变成 LF，不先拆段后清理两侧空格。[PlcfSed](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/68a959a6-11e0-4f3e-9a99-76ca8cc4dddc)。
2. 找到段落结尾对应属性，读取 PAPX，再按规范应用结尾 PCD 的 PRM 修改。不能假设首个 piece 的属性代表整个段落。[直接段落属性](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/61b635c3-2c44-4155-bf17-fec281b30c71)。
3. 至少识别 `PFInTable`、`PFTtp`、`PItap`、`PDtap`、`PFInnerTableCell`、`PFInnerTtp`。`PDtap` 是有顺序的层级增量，不能把属性简单存入字典后只保留最后一个值。[段落属性](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/484822ee-a9d9-4af4-8423-29fda67a6a58)。
4. 支持 Prm0 的相关简短属性及 Prm1 指向的 CLX 属性组。遇到 `PHugePapx`、`PTableProps` 等间接属性时，按规范读取 Data stream；遵守替换和应用顺序，限制引用深度、循环和总处理量。
5. 跳过不需要的属性时，仍按 SPRM 的真实操作数长度前进，处理可变长度例外，不能按固定长度猜下一条。PAPX 字节数按 `PapxInFkp` 解释，不无条件吞掉末尾孤立字节。[Sprm](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/099eb99c-a927-4caf-a80c-66254ea83d6a)、[PapxInFkp](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/580510b8-df7a-467e-a51c-0d71eb15c7cd)。

内部使用段落、单元格、行的小型记录保存这些关系。

DOC 的物理单元格边界不能直接视为 DOCX 的 `tc` 数量。平台对直接 `tc` 逐个输出，对 `gridSpan`、`vMerge` 不展开、不复制文字。

| 情形 | 平台基准 | 对 DOC 实现的要求 |
| --- | --- | --- |
| 普通两格 | `A\tB` | 依据行、格结束建立两个输出格 |
| 一个 `tc` 跨两列，后跟另一个 `tc` | `A\tB` | 不为跨列跨度补空格；需确认 DOC 的合并记录如何对应这个 `tc` |
| 垂直合并首行为 A/B，次行续格为空/C | `A\tB\n\tC` | 不复制上方 A；保留前置空格子的分隔，随后按拼接规则清理行尾 |
| 一个 `tc` 含嵌套表格 | 后代段落用 LF 拼接 | 内层单元格不产生外层 TAB，不额外输出内层行结束标记 |

需要读取表格定义及适用的合并属性，而不是仅靠 `0x07`。规范给出了 `TDefTable`、`TInsert`、`TDelete`、`TMerge` 与单元格属性；实际 DOC 记录的单元格、合并区域和最终输出格需要明确映射。[表格属性](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/b39a6648-501c-4361-8366-4f042f579469)、[TDefTableOperand](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/de06ec41-a0ac-4046-9096-cdfaa0091ad9)、[TC80](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9dd62a79-8c0b-4b11-99ee-05742ae7cf6d)。

跨格式表示本身可能不同：DOCX 也可能以多个带合并标记的 `tc` 表示一个区域，而平台会照读这些 `tc`。因此“画面一样”不够作为逐字相同的验收条件。对每种声称支持的映射，必须提供结构明确的配对 fixture，不能由某个转换器输出反推唯一规则。

已核对的 DOC 细节：`TCGRF.horzMerge=1` 表示横向合并续格，`2/3` 表示首格；`sprmTMerge` 也可给出横向合并。纵向合并的续格仍计入物理格索引，最终状态还可能由 `sprmTVertMerge` 覆盖。解析须应用有效行属性，不能只读取初始 TC80。[TCGRF](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/11bf5b1c-943f-421d-bbf3-39088cd1b8dd)、[VertMergeOperand](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/cf7489ac-7eee-404d-8844-8ebbd279b77d)。

规范中“续格的内容不渲染”不能单独作为删除文字的依据，因为平台也不按可见性过滤 DOCX 文字。遇到续格含字的样例，应先确认其合法性和对应的 DOCX 存储，再确定输出；不能直接将所有续格内容强制清空。上述垂直合并基准只针对内容为空的续格。

嵌套表格最终按所属单元格内的后代段落展开，不输出独立的内层列分隔。仍须维护嵌套层级，以免把内表结束误当作外表结束。

### 3.4 按固定规则拼接文字

对已建立对应关系的 DOC 内容，输出按以下顺序构造：

1. **段落内文字。** 按读取顺序拼接文字、字段代码和结果、修订文字；TAB 保留为 `\t`，手动换行对应 `\n`。平台把 `br` 和 `cr` 都当换行，包括有 page/column 类型的 `br`。DOC 中分页、分节及对象特殊标记须先按格式识别，不能把所有控制字符都当普通换行。
2. **清理单段文字。** 使用 `text.replace("\r\n", "\n").replace("\r", "\n").strip(" \n")`。只去首尾空格和 LF；保留内部连续空格、TAB，以及没有被对应 XML 读取规则排除的普通 Unicode 字符。
3. **单元格。** 按对应的后代段落顺序，过滤清理后为空的段落，用 LF 拼接剩余段落。嵌套表格按段落展开。
4. **顶层表格行。** 单元格用 TAB 拼接，再调用不带参数的 `.rstrip()`。它去除行末所有 Python 空白字符，包括末尾空单元格留下的 TAB；不清理行首 TAB。
5. **整篇正文。** 顶层段落和顶层表格行按顺序用 LF 拼接。中间空段落、空行保留位置；最后再执行第 2 步的 `_clean_text()`。

DOCX body 的其他包裹节点走的是“枚举后代段落”分支，不会再次套用顶层表格行列序列化。DOC 若有可对应的内容容器，必须先确认结构对应，不能假定所有表格都天然获得顶层 TAB 规则。

基线 `normalize.py` 的全局正则会合并连续空格和 TAB、删除空行、删除 PAGE/NUMPAGES 字段文字、过滤私用区字符及 U+FEFF。这与基准不一致。新的对齐输出路径需要移除这些清洗步骤，以选定的内容范围和结构记录决定输出。

例如字面正文 `A  B` 的两个空格应保留；一个实际字段中存储的 ` PAGE ` 和结果 `3`，在该段首尾清理后是 `PAGE 3`。不会为了让结果更像 Word 画面而自行隐藏字段代码或修订删除文字。

DOC 结构控制符需要消费，而真实文字需要保留。不能为了停用旧清洗就把 `0x07`、字段边界、对象锚点直接写进文本，也不能继续按全局字符黑名单删除本可对应到 XML 文本的字符。

### 3.5 区分格式映射与视觉相同

可直接比较的是明确对应的内容和结构。对同样的主体段落、表格和字段构造各自合法的 DOC、DOCX 样例，先确定 DOCX 基准文本，再检查 DOC 输出。不要把某个转换器生成的所有 XML 细节当作 DOC 规范保证。

平台遍历还存在两种必须单独记录的行为：

- 正文段落内的文本框有两段“框一”“框二”时，递归 inline 读取会得到 `前框一框二后`，不会为框内段落自动加换行。
- 同样的文本框放在表格单元格中，父段落和后代段落都会被读取，可能得到 `前框一框二后\n框一\n框二`。这是已验证的遍历重复，不能把“保持可见文字一致”与“逐字复制当前输出”混为同一个承诺。

目前的决策是沿用平台逻辑，因此文档保留这些基准结果，不在此次 DOC 设计中顺带修正平台。如何把 DOC 的锚点和独立 story 投影成相同的遍历关系，仍需合法样例验证；未通过前，不能宣称文本框完全对齐。

艺术字也是表示形式的边界：DOC 的 OfficeArt 可以把文字存入 `gtextUNICODE_complex`，而 DOCX 的文字可能成为元素文本，也可能成为 VML 属性。相同画面不保证相同基准输出。确定实现范围前需记录实际对应形式；不能自行扩大为“解析所有 OfficeArt 字符串”，也不能把可读取文本框和艺术字一概排除。[gtextUNICODE](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/3b9a1a7b-ad49-410c-bb32-ab91decff765)、[实际 Unicode 字符串](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/c2dfbfd1-9652-40c2-a8d2-1f292bae61b5)。

## 4. 损坏、兼容与资源边界

| 情形 | 结果 |
| --- | --- |
| 所选内容按规则拼接后为空，所需结构合法 | 成功返回 `text=""` |
| FIB 指定的 CLX 缺失、越界或结构不成立 | 整体失败；不扫描其他候选记录回退 |
| 必需 piece、字符范围或编码损坏 | 整体失败，不静默跳过，不以替换字符掩盖已确认的文字损坏 |
| 已选正文或已引用文本框的必要关联、段落或表格结构损坏 | 整体失败，不返回文字或顺序不完整的成功结果 |
| 未选 story 的独立内容损坏 | 不主动解码或深入验证；确定所选内容所需的共享记录仍须合法 |
| 可选标题、作者等元数据损坏 | 正文可读时继续，沿用 metadata warning |
| 加密、不支持的版本、资源超限 | 给出可区分的错误原因，不一律称为文件损坏 |

严格读取不等于实现完整 DOC 验证器。校验所选内容依赖的结构，尊重必须忽略的保留字段。规范允许的写入差异应在同一路径内处理，不通过恢复旧扫描兼容。过去靠扫描碰巧读出的文件可能报错，这是已接受的兼容代价。

Word 提供恢复损坏文件的路径，因此“某个应用能打开”只能作为兼容线索，不能证明所有记录符合规范；库不承担文件修复。[微软文件修复说明](https://support.microsoft.com/en-US/Word/open-a-document-after-a-file-corruption-error)。

保留现有最大文件、stream 和输出限制。读取及分配前检查偏移、长度、计数和加法边界；按增量检查输出 UTF-8 大小，不能只在完整大字符串生成后检查。对重复 piece 引用、文本框引用、嵌套表格和间接属性设置有界处理，防止异常工作量及循环。文字结构预算不能借用语义无关的 OLE 链长度选项；具体内部阈值仍需依据规范和样例确定。

单元格结束、行结束和嵌套层级须相互一致；检查行格数、合并范围时应读取相应表格定义。不能仅凭基础层级标记就宣称完成全部表格一致性校验。

## 5. 源码修改与实施顺序

| 位置 | 改动职责 |
| --- | --- |
| `src/legacy_doc/word.py` 及必要私有解析模块 | FIB/CLX 定位、CP/FC 映射、所选 story 和文本框引用；复用 OLE reader |
| 私有段落与表格解析代码 | PAPX、PRM、间接属性、嵌套层级、行格及合并关系 |
| `src/legacy_doc/normalize.py` 或私有文字拼接代码 | 消费结构控制符，执行段落、格、行的确定规则；移除旧正则压缩路径 |
| `src/legacy_doc/api.py`、`word.py` | 默认调用新路径，合法空正文正常返回，落实增量输出限制 |
| `src/legacy_doc/types.py` | 沿用现有结果和选项；不新增旧输出模式或公开表格对象 |
| `tests/fixtures.py` 与针对性测试 | 合法 FIB、CLX、cbMac、UTF-16 CP、多 piece 和必要属性构造 |
| README、发布说明 | 写明默认输出变化、提取范围、空结果与严格失败规则 |

先补合法 fixture 和预期值，再完成定位与范围读取、段落及表格关系、最终拼接和公开入口切换。私有模块名称与拆分以实际代码职责决定，不以文件数量作为目标。字体、颜色、视觉布局、分页计算、OCR、文件转换和完整 Word 对象模型不属于本次实现。

## 6. 验收标准与已有证据

验收区分三类证据：旧库回归证明原有能力未意外破坏；DOCX 基准样例证明输出规则理解正确；合法 DOC fixture 才能证明新解析实现了对应关系。期望值事先确定，不以另一个应用能打开或某个转换器输出作为唯一标准。

### 6.1 文字定位与完整性

| 编号 | 构造或输入 | 必须成立的结果 |
| --- | --- | --- |
| T01 | 0Table/1Table 各一份合法文件 | 均从 FIB 指定的 Table stream 正确取字 |
| T02 | 真 CLX 前面放一段形状像 piece table 的数据 | 仍读取 FIB 指定 CLX，不受假候选影响 |
| T03 | 指定 CLX 已损坏，文件别处有貌似可用候选 | 整体失败，不读取替代候选 |
| T04 | 三段正文，中间 piece 越界 | 整体失败，不返回第一段和第三段组成的残缺文本 |
| T05 | PCD 指向 cbMac 之后、实际 stream 结束之前 | 拒绝该无效范围 |
| T06 | 改变必须忽略的 0x18/0x1C 保留字段，其他记录保持合法 | 不影响正确定位 |
| T07 | 中文、英文和压缩/非压缩 piece 混排，FC 次序不同于 CP 次序 | 输出按文档顺序，字符完整 |
| T08 | `A😀Z`，以及代理对跨两个合法 piece | 保留 emoji 和末尾 Z；fixture 本身按 UTF-16 单元计数 |
| T09 | 请求区间从一个 piece 中间开始，到另一个 piece 中间结束 | 只读取指定 CP 区间；不多取、不漏取 |
| T10 | FIB/CLX 截断、错误计数、不支持版本或解码必需字节不足 | 明确失败，错误原因可区分 |
| T11 | 正文合法，SummaryInformation 损坏 | 文字照常返回，metadata 附 warning |
| T12 | 重复片段引用等导致异常工作量或资源超限 | 有界失败，不进行无界分配或循环 |

### 6.2 内容范围与特殊文字

| 编号 | 验收内容 |
| --- | --- |
| S01 | 正文、页眉页脚、脚注、尾注、批注各放唯一标记，仅正文标记输出 |
| S02 | 主体范围跨 piece，尾部紧接另一个 story，既不越界也不漏最后一个正文字符 |
| S03 | 正文文本框、页眉文本框、闲置文本框分别放不同文字，只访问正确的正文关联 |
| S04 | 文本框位于正文中间、表格单元格以及连续链接框中，检验插入位置、遍历次数和边界 |
| S05 | 字段代码、字段结果、修订删除文字和字面 PAGE 字样保留；结构控制符不泄漏 |
| S06 | 已引用文本框的记录损坏，整体失败；未选 story 的独立内容损坏不自动导致失败 |
| S07 | 艺术字按实际存储形式建立配对样例，逐项记录能对应、不能对应及其原因 |

### 6.3 段落、表格和属性覆盖

| 编号 | 验收内容 |
| --- | --- |
| R01 | 普通段落、两行两列表格、后续普通段落，按文档顺序输出 |
| R02 | 同一段落跨 piece，压缩与非压缩 FC 均正确定位到段落属性 |
| R03 | 仅由 Prm0、Prm1、PAPX 或间接属性给出关键标记，分别验证实际生效 |
| R04 | 重复 PDtap、PAPX 与 PRM 覆盖，验证有序应用 |
| R05 | 单元格多段落、空格子、空行、两层及三层嵌套，内外层边界不混淆 |
| R06 | 水平合并、垂直合并及组合合并，按选定存储对应关系与 DOCX 基准比较 |
| R07 | 相邻行格数不同、相邻独立表格，文字顺序及分隔不依赖矩形假设 |
| R08 | FKP 越界、属性截断、非法层级、间接引用循环或超限，明确失败 |

### 6.4 默认返回与输出样例

- 不传新增选项的 `extract_text()` 直接返回新规则结果；没有切回旧输出或旧定位的模式。
- 合法空正文返回现有结果类型，`text=""`，不抛“未提取到文字”异常；仍返回可用元数据。
- 损坏、加密和超限仍按错误处理，不能通过空字符串掩盖。
- 所选字段、修订及隐藏文字保留，结构控制符不泄漏；普通 Unicode、连续空格及跨 piece 的有效代理对完整。

下表用 `\t`、`\n` 表示真实 TAB 和换行。15 个合成 XML ZIP 已通过隔离执行参考函数验证；完整输入、源码哈希和结果见 [docx-baseline-cases.json](docx-baseline-cases.json)。这些是提取函数行为样例，不是已通过 Word 格式验证的文件，也不是新 DOC 实现通过记录。

| 输入结构 | 基准输出 |
| --- | --- |
| 段落 `  A  B  `、空段落、段落 C | `A  B\n\nC` |
| 一行五格：空、B、空、D、空 | `\tB\t\tD` |
| 左格段落 A、空、B；右格 C | `A\nB\tC` |
| 左格包含“前”、内表两格甲/乙、“后”；右格“右” | `前\n甲\n乙\n后\t右` |
| 跨两列的一个 `tc` 为 A，下一个 `tc` 为 B | `A\tB` |
| 字段代码 PAGE、结果 3、删除文字旧、插入文字新 | `PAGE 3旧新` |
| 正文内文本框，两段为框一/框二 | `前框一框二后` |
| 相同文本框位于单元格中 | `前框一框二后\n框一\n框二` |
| 内容控件包裹两格表格 A/B | `A\nB` |
| `a:t` 为节点字，VML 属性为属性字 | `节点字` |

旧规范化函数的两行表格样例输出为 `姓名 年龄 张三 28`；新规则目标为 `姓名\t年龄\n张三\t28`。主体 `第一段\r\r第二段\r` 原先丢失中间空段，新规则保留为 `第一段\n\n第二段`。

旧库基线回归曾有 21 项通过。项目的整套测试结果为 **139 passed**；其中包括旧 API/元数据回归、二进制记录测试、表格和文本框整文件测试。完整验证命令与样本结果见 [source-checks.json](source-checks.json)。这些证据覆盖本次构造和已保存样本，不等于通过所有 DOC 文件的兼容性认证。

## 7. 当前实现与证据边界

### 7.1 实现位置

独立项目：`/Users/gearup/workspaces/legacy-doc-development/doc-text`，分支 `main`，有自己的 `.git`，没有远端。发行包名及 parser 标识为 `doc-text`，独立版本从 `0.1.0` 开始；Python 导入仍为 `legacy_doc`。来源为 legacy-doc 提交 `33aac7cace0dc30d95e671df5aaa825e4c2223a1`，保留原许可与来源说明。旧二开 worktree 在文件哈希核对和迁移验证后已移除；上游参考仓库仍保留。

| 私有模块 | 当前职责 |
| --- | --- |
| `_binary.py` | FIB 有效版本、指定 CLX、CP/FC、Prc/PRM、严格解码及 Data 按需读取 |
| `_paragraphs.py` | PAPX/PRM、间接属性、表格行格与合并标记、PlcfSed 分节识别 |
| `_shapes.py` | 正文 SPA、OfficeArt、FTXBXS、链接框范围和关联完整性 |
| `_characters.py` | 按需读取 CHPX 与 PRM，区分符号占位和普通括号，验证文本框锚点的 CFSpec |
| `_text.py` | 沿用选定 DOCX 的段落、表格、嵌套和文本框遍历；检查输出与工作量 |
| `word.py`、`normalize.py` | 所有现有入口使用新路径，元数据 warning，合法空正文；移除旧扫描和正则删文 |

公开的 `extract_text()`、`ExtractionOptions` 和 `DocExtractionResult` 沿用原接口，没有新增旧输出模式。README 记录了从本地 checkout 安装的方法和默认输出变化。Python 运行时依赖仍为空。

### 7.2 验证结果

- 整套测试：**139 passed**。200 份真实 DOC 的批量执行与结果分类见第 7.5 节；逐字输出断言仅覆盖有明确预期的回归样本。
- 普通表格、合并表格、嵌套表格及受控多语言正文四份已有 DOC 的输出，与保存的 DOCX 基准逐字一致；没有重跑转换。
- 仓库包含一份本地自编 RTF 导出的真实 DOC，用于回归中文、非 BMP 字符、空单元格、多段落格和正文边界。
- 正文文本框、单元格中的文本框及连续链接框使用完整合成 OLE 文件验证插入位置与遍历次数；关联损坏、链条缺段等不得返回残缺成功结果。
- `testPictures.doc` 可读取，但与保存的转换后 DOCX 文本仍有差异。该样本的 DOC 保留字段代码等内容；转换后的表示不同，不能作为任意 DOC/DOCX 等价的证明。

### 7.3 明确的边界

1. **真实文本框覆盖有限。** 已加入 LibreOffice 上游的两份正文文本框 DOC，文件元数据均标记 Microsoft Office Word；分别验证透明文本框文字和六种填充文本框中的文字。WPS 及真实连续链接框仍未形成专门样本覆盖。链接框当前要求分段和锚点完整覆盖；规范允许但只提供部分范围的记录会明确拒绝，以免返回漏段结果。
2. **艺术字表示不唯一。** 已选正文形状若只以 OfficeArt geometry-text 属性保存文字，当前明确报不支持，避免静默漏掉；不擅自将所有 OfficeArt 字符串当正文。
3. **样式相关特殊字符。** CFSpec 的 `0/1` 及 CHPX、Prm0、Prm1 覆盖已处理。最终值若仍依赖 `0x80/0x81` 的样式求值，且提取需要判定该字符，则明确报不支持；本轮未实现完整 Word 样式继承。
4. **表格与包裹结构。** 当前投影保留已存文字并使用选定的水平合并、纵向续格和嵌套规则；不承诺任意内容控件、图形版式或同画面的 DOC/DOCX 都逐字相同。

平台接入、提示浮窗、空正文提示文案和发布流程仍不在本轮范围。后续样本补充应围绕上述证据缺口，不重开已经确认的默认接口、空正文和提取规则。


### 7.4 借用上游样本并改写验证条件

沿用用户确认的最小提取范围。上游项目仅提供 DOC 文件和测试条件，按我们的分类改写断言；不引入其他语言的解析器或转换运行时。当前仅本地自用，不推进上游 PR 或发布。

| 我们的分类 | 样本数 | 改写后的条件 |
| --- | --- | --- |
| 正文文字 | 1 | 保留段落文字、内部空段，统一 LF |
| 正文范围 | 2 | 保留正文 Unicode；排除页眉、页脚、脚注、尾注和批注 |
| 空正文 | 1 | 返回空串、零文字长度，正常成功 |
| 表格 | 3 | 普通表格、合并格、嵌套表格符合既定文字顺序与 TAB/LF 规则 |
| 正文文本框 | 2 | 透明或不同填充属性均不影响框内文字提取 |
| 不支持的格式 | 1 | Word 95 明确失败，保持现有支持范围 |

文件、固定来源提交、SHA-256、上游测试函数和条件映射记录在 项目的 `tests/data/upstream/cases.json`。执行入口为 `tests/test_real_documents.py`，不访问网络、不调用 Office。原始文件共 305,152 字节，未修改；保留来源许可说明，不把第三方样本重新标为本项目 MIT。

两份文本框样本暴露并修复了一个边界检查问题：`PlcfSpaMom` 的最终 CP，以及 `PlcftxbxTxt`、`PlcfTxbxBkd` 最后未使用区间的结束 CP，可以超出所选 story 范围。原实现错误地用这些值判断越界。现在忽略不参与文字读取的末尾值，实际锚点、有效文字区间和 CP 严格递增仍受检查。[FibRgFcLcb97](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/0c9df81f-98d0-454e-ad84-b612cd05b1a4)、[PlcftxbxTxt](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/9f17f89a-64d6-4546-a811-e62e1d238102)、[Tbkd](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/29b6f46c-7136-4a7b-9de4-6f17a6e4e677)。

验证：新增末尾 CP 正例先出现 3 项失败；修复后整套 **116 passed**，ruff F 和 diff 空白检查通过。旧构建验证对应此前版本，本轮记录以 `source-checks.json` 的 `borrowed_fixture_validation` 为准。


### 7.5 独立项目与 200 份 DOC 批量验证

项目已迁至 `doc-text/`，有自己的 Git 仓库和 `main` 分支。发行包、结果 parser 标识为 `doc-text`，版本 `0.1.0`；导入保持 `legacy_doc`。原 48 个项目文件迁移前后逐一核对 SHA-256。旧 worktree 已移除，上游参考仓库保留。唯一方案和验证记录随项目置于 `docs/`，后续仅在此项目维护。

语料清单为 `tests/data/corpus.json`：200 份 SHA-256 唯一文件，共 23,630,630 字节，其中 Apache POI 156 份、LibreOffice 44 份，复用此前 10 份样本。固定上游提交、Git blob SHA 和下载文件哈希均已核验。49 份新增样本有核实后的上游条件，10 份沿用已有精选条件，5 份分类仍标记不确定，其余按文件名信号或通用回归用途记录；不把通用回归当作逐字正确性的证明。已核实分类保存在 `tests/data/corpus-conditions.json`，抓取脚本会重新应用，避免下载刷新时丢失。

本轮批量结果：148 份正常返回文字；23 份不支持的 Word 版本、3 份加密、1 份不支持的容器特性、1 份非 OLE 输入；5 份触发资源上限；19 份结构拒绝仍需逐项核对。未捕获异常和超时均为 0。19 份中包含上游损坏回归文件，也包含段落、表格等解析差异，不能统一宣称为无效文件。完整错误和来源见 `docs/corpus-results.json`。

相较首轮 142 份正常返回、1 份未捕获异常，已完成：

- OLE 根存储按对象类型识别，允许真实样本使用非 `Root Entry` 的名称，恢复 4 份文件。
- FIB 先读取有效版本，再核对已知布局；已知扩展前缀后的额外数据按声明长度有界跳过，恢复 2 份文件。短字段、截断、未知版本仍拒绝。
- 无法表示的 FILETIME 元数据日期转为既有元数据 warning，避免 OverflowError 中断处理。
- 通过 `0xA5DC` 标识报告不支持的 Word 6/95，而非把旧格式一律标为 FIB 损坏。

验证命令：`python3 scripts/check_corpus.py`（有待核对拒绝时退出码为 1）；pytest **139 passed**；ruff F 通过；独立 wheel 构建通过。仅本地保存，没有远端、PR 或发布。
