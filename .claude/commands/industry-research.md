---
description: 对指定行业进行系统性入门研究，输出包含产业链地图、竞争格局、景气周期、政策环境的完整行业认知报告
argument-hint: "[行业名，如：PCB、光模块、风电、汽车零部件、消费电子供应链、石英频控、精密电机、AI芯片、激光设备、农牧养殖]"
---

# 行业入门研究命令

## 使用示例

```
/industry-research PCB
/industry-research 光模块
/industry-research 精密电机
/industry-research 风电
```

## 执行流程

### Step 1：解析行业名称

从参数识别行业，匹配行业索引：

| 输入关键词 | 对应行业文件 |
|-----------|-----------|
| PCB / 印制电路板 | industries/PCB.md |
| 光模块 / AI算力 / 数通 | industries/光模块_AI基础设施.md |
| 风电 / 风机 | industries/风电.md |
| 新能源汽车 / 电动车 | industries/新能源汽车.md |
| 汽车零部件 / 铝压铸 / 底盘 | industries/汽车零部件.md |
| 消费电子 / 天线 / FPC | industries/消费电子供应链.md |
| 石英 / 晶振 / 频控 | industries/石英频控.md |
| 步进电机 / 伺服 / 运动控制 | industries/精密电机.md |
| AI芯片 / GPU / 算力芯片 | industries/AI芯片.md |
| 激光 / 激光设备 | industries/激光设备.md |
| 农牧 / 猪 / 养殖 | industries/农牧养殖.md |
| 商业航天 / 火箭 / 卫星 / 星座 | industries/商业航天.md |
| 人形机器人 / 具身智能 / 机器人 | industries/人形机器人.md |
| 核聚变 / 托卡马克 / 聚变能源 | industries/核聚变.md |
| 无人驾驶 / 自动驾驶 / Robotaxi / 智驾 | industries/无人驾驶.md |

若输入不在列表，询问用户最接近的行业，或提示"暂无该行业模板，是否启用通用框架？"

### Step 2：加载 Skill 和行业文件

```
Read: d:\claude-for-finance\.claude\skills\industry-research\SKILL.md
Read: d:\claude-for-finance\.claude\skills\industry-research\industries\[对应文件]
```

### Step 3：执行五步行业入门研究

使用 WebSearch 按行业文件中的 Prompt 依次查询：
1. 产业链结构
2. 竞争格局
3. 景气周期与领先指标
4. 政策环境
5. 关键财务特征

每步查询后整合结果，不要等所有步骤完成再输出。

### Step 4：输出研究报告

报告格式：
```markdown
# [行业名] 行业研究笔记
**研究日期**：[日期]
**关联标的**：[watchlist中对应标的]

## 一、产业链地图
## 二、竞争格局
## 三、景气周期与领先指标
## 四、政策环境
## 五、财务特征与估值参考
## 六、近期核心关注点
```

保存至：`d:\claude-for-finance\outputs\行业研究\[行业名]-入门-[YYYYMMDD].md`

若目录不存在，自动创建。
