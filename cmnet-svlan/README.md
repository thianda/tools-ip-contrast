## 基本思路

1. 采集各 bras 的 `vpn-instance LNTL_PON_Manager_HW`和`vpn-instance LNTL_PON_Manager_ZTE`的`ip routing-table`信息
2. 基于`ip`匹配 OLT 名称与`bras`、`Trunk`口信息的关系
3. 梳理现网`Trunk`口的子接口数据
4. 自动生成新增`svlan`脚本 