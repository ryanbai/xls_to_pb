
# 说明

---

游戏的配置数据常常是在excel里完成的。将excel中的数据导出给程序员使用，需要定义转换和解析接口这些手工活，这个工具是为了节省人力，实现最大程度的自动化。

这个工具的灵感来自jameyli同学的导表工具，他们把proto的定义部分放在excel的表头，这样就不用人工维护一个和这张表对应的proto。在此基础上，我增加了公共数据结构的导入、protobuf保序等工作。

算法的核心思想是一个sheet页对应一个pb的message（可以看成一棵树）定义，sheet页的每一列都是message的根节点到叶子节点的路径，那么构建message的描述的过程就是恢复这棵树的过程。

# 用法

---

用法比较简单，指定sheet名和excel名：
```shell
./xls_pb_tool.py ITEM_CONF 道具表.xls
```

excel文件：


生成的proto文件：


生成的text文件：

