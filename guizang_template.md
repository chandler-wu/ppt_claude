使用 guizang-ppt-skill 生成一份线下分享 deck。
素材：【你的大纲 /markdown 文件/ 提示词输入】
风格：Style A 电子杂志 / Style B 瑞士国际主义
主题色：【五选一 / 四选一】
结构原则：Hero 开场 + 收尾必须有，Hero / 内页交替
输出单文件 HTML。

示例
- 场景名 → Q2 业务复盘汇报
- 主题 → `corporate-clean`（商务简洁风）
- 模板 → `weekly-report`（周报/月报 full-deck）
- 页数 → 10-12 页
- 页脚 → "XX 部门 · 2026 Q2"
- 特殊要求 → "必含 KPI 大字报 + 业务趋势折线图 + Q3 规划 roadmap"

演讲者模式示例
用 guizang-ppt-skill 生成了一份 deck。
现在想加上 html-ppt 的演讲者模式。
请完成：
读取 html-ppt 演讲者模式源码
移植到歸藏 deck：每页加 <aside class="notes">
追加演讲者视图 JS（S 键弹窗）
保留歸藏原有视觉不破坏