# slide09_whitepaper_page09_interactive

- Source: `slide09_whitepaper_page09_interactive.pptx`
- Total slides: 1

## Slide 1

09

结果分析

控制变量法下的摩擦系数变化规律 / Controlled Variable Analysis of Metal-Ice Friction Coefficient

六因素模型框架

点击节点查看变量分析态

温度 / Temperature

压强 / Pressure

金属材质 / Metal Material

表面粗糙度 / Surface Roughness

接触面积 / Contact Area

冰的成分 / Ice Composition

F01

F02

F03

F04

F05

F06

TRIGGER

TRIGGER

TRIGGER

TRIGGER

TRIGGER

TRIGGER

实验手段选择与模型构建

F01

F02

实验水平表

实验水平表

实验水平表

实验水平表

实验水平表

实验水平表

摩擦系数 μ 趋势

摩擦系数 μ 趋势

摩擦系数 μ 趋势

摩擦系数 μ 趋势

摩擦系数 μ 趋势

摩擦系数 μ 趋势

℃

P

以控制变量法组织实验水平，将装置设计、变量设计与数据处理串联为可替换的数据分析入口。

> [Chart] chart_group_temperature — LINE_MARKERS (65)

| Category | μ |
| --- | --- |
| -15 ℃ | 0.118 |
| -10 ℃ | 0.103 |
| -5 ℃ | 0.089 |
| 0 ℃ | 0.076 |

> [Chart] chart_group_pressure — LINE_MARKERS (65)

| Category | μ |
| --- | --- |
| 0.05 MPa | 0.082 |
| 0.10 MPa | 0.096 |
| 0.20 MPa | 0.091 |

> [Chart] chart_group_material — COLUMN_CLUSTERED (51)

| Category | μ |
| --- | --- |
| 铝合金 | 0.108 |
| 不锈钢 | 0.095 |
| 钛合金 | 0.083 |

> [Chart] chart_group_roughness — LINE_MARKERS (65)

| Category | μ |
| --- | --- |
| Ra 0.2 μm | 0.074 |
| Ra 1.0 μm | 0.096 |
| Ra 3.0 μm | 0.124 |

> [Chart] chart_group_area — LINE_MARKERS (65)

| Category | μ |
| --- | --- |
| 1 cm² | 0.112 |
| 4 cm² | 0.096 |
| 9 cm² | 0.087 |

> [Chart] chart_group_ice — COLUMN_CLUSTERED (51)

| Category | μ |
| --- | --- |
| 淡水冰 | 0.104 |
| 海冰 | 0.089 |
| 人工盐冰 | 0.078 |

温度

压强

| 水平 | 变量取值 | μ |
| --- | --- | --- |
| T1 | -15 ℃ | 0.118 |
| T2 | -10 ℃ | 0.103 |
| T3 | -5 ℃ | 0.089 |
| T4 | 0 ℃ | 0.076 |

| 水平 | 变量取值 | μ |
| --- | --- | --- |
| P1 | 0.05 MPa | 0.082 |
| P2 | 0.10 MPa | 0.096 |
| P3 | 0.20 MPa | 0.091 |

| 水平 | 变量取值 | μ |
| --- | --- | --- |
| M1 | 铝合金 | 0.108 |
| M2 | 不锈钢 | 0.095 |
| M3 | 钛合金 | 0.083 |

| 水平 | 变量取值 | μ |
| --- | --- | --- |
| R1 | Ra 0.2 μm | 0.074 |
| R2 | Ra 1.0 μm | 0.096 |
| R3 | Ra 3.0 μm | 0.124 |

| 水平 | 变量取值 | μ |
| --- | --- | --- |
| A1 | 1 cm² | 0.112 |
| A2 | 4 cm² | 0.096 |
| A3 | 9 cm² | 0.087 |

| 水平 | 变量取值 | μ |
| --- | --- | --- |
| I1 | 淡水冰 | 0.104 |
| I2 | 海冰 | 0.089 |
| I3 | 人工盐冰 | 0.078 |

Temperature

Pressure

T1 · T2 · T3 · T4

P1 · P2 · P3

金属-冰接触实验场景

F03

F04

▰

≈

μ

金属材质

表面粗糙度

Metal Material

Surface Roughness

- 变量：温度
- 因变量：摩擦系数 μ

- 变量：压强
- 因变量：摩擦系数 μ

- 变量：金属材质
- 因变量：摩擦系数 μ

- 变量：表面粗糙度
- 因变量：摩擦系数 μ

- 变量：接触面积
- 因变量：摩擦系数 μ

- 变量：冰的成分
- 因变量：摩擦系数 μ

M1 · M2 · M3

R1 · R2 · R3

→

→

→

→

控制变量法

实验水平设计

μ 数据采集

趋势建模

结果分析

温度升高时，冰表面液膜增强，摩擦系数 μ 整体下降。

压强增大改变局部接触状态与压力熔化程度，摩擦系数可能呈现先升后降趋势。

不同金属材料的导热性、表面能和微观接触状态不同，导致摩擦系数存在显著差异。

表面粗糙度增大时，微观嵌入和犁削效应增强，摩擦系数 μ 上升。

在其他条件不变时，接触面积变化会改变单位压强和局部液膜状态，使摩擦系数发生变化。

冰中盐分或杂质会改变冰点与表面液膜状态，从而影响金属-冰界面的摩擦行为。

F05

F06

□

⌬

接触面积

冰的成分

Contact Area

Ice Composition

初始态保留模型总览；节点触发后切换为实验水平表、趋势图与物理结论。

A1 · A2 · A3

I1 · I2 · I3

注：当前数据为版式示例，后续可替换为真实实验数据。

09
