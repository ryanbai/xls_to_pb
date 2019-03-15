#! /usr/bin/env python
#coding=utf-8

##
# @file:   xls_pb_tool.py
# @author: ryanbai
# @brief:  xls 配置导表工具
# @desc:  参考jameyli的xls_deploy_tool.py工具
#
# 主要功能：
#     1 配置定义生成，根据excel的内容自动生成配置的PB定义
#     2 配置数据导入，根据PB将配置数据序列化为二进制数据或者文本数据
#
# 设计思想：
#    如果把一个message的pb描述看作一颗树，那么将根节点到叶节点的路径映射到excel的每一列,
# 读取excel的每一列，根据路径就可以重建这棵树。
#
# 依赖:
# 1 protobuf
# 2 xlrd
##
import sys
import json
import os
import traceback
import xlrd

# utf-8编码方式
reload(sys)
sys.setdefaultencoding( "utf-8" )

# 创建文件夹
def create_path_if_noexist(pathname):
    """目录如果不存在，创建"""
    if os.path.isdir(pathname):
        return
    os.mkdir(pathname)

# common_def文件定义依赖的外部变量
# PROTOC: protoc的路径
# PB_PROTO: pb自带的proto文件
# PB_PYTHON: pb的python库
# CUSTOM_PROTO: 项目的proto目录
# CUSTOM_PYTHON: 项目的proto生成的python库目录
# COMMON_MODODULES: 依赖的公共模块。这些模块中定义的message可以直接使用。
from common_def import *
sys.path.append(PB_PYTHON)

# 本地python路径 
sys.path.append(PYTHON_GEN_PATH) 

# 加载公共库
sys.path.append(CUSTOM_PYTHON)
comm_loaded_modules = map(__import__, COMMON_MODULES)

def decimal2az(n):
    """将十进制转换为26进制a-z"""
    return (decimal2az(n/26-1) if n/26>0 else '') + chr(65 + n%26)

#########################protobuf cpp_type##########################################
#CPPTYPE_INT32       = 1,     // TYPE_INT32, TYPE_SINT32, TYPE_SFIXED32
#CPPTYPE_INT64       = 2,     // TYPE_INT64, TYPE_SINT64, TYPE_SFIXED64
#CPPTYPE_UINT32      = 3,     // TYPE_UINT32, TYPE_FIXED32
#CPPTYPE_UINT64      = 4,     // TYPE_UINT64, TYPE_FIXED64
#CPPTYPE_DOUBLE      = 5,     // TYPE_DOUBLE
#CPPTYPE_FLOAT       = 6,     // TYPE_FLOAT
#CPPTYPE_BOOL        = 7,     // TYPE_BOOL
#CPPTYPE_ENUM        = 8,     // TYPE_ENUM
#CPPTYPE_STRING      = 9,     // TYPE_STRING, TYPE_BYTES
#CPPTYPE_MESSAGE     = 10,    // TYPE_MESSAGE, TYPE_GROUP
##############################################################################
# 描述树
class DescTree:
    """描述树的节点"""
    # 最近遍历过的叶节点
    __last_leaf_node = None
    # 用到的公共模块
    __used_modules = set()

    def __init__(self, name, type_name, comment, pb_desc = None):
            """构造节点时指定名称"""
            self._name = name
            # 类型
            self.__init_private_type(type_name)
            # 读取第3行做注释
            self._comment = comment
            # 是否是公共类型
            self._is_common_type = False
            # 当前节点重复的次数
            self._repeated_count = 1
            # 子节点
            self._sub_nodes = []
            # 父节点
            self._parent = None
            # 当前节点的深度
            self._depth = 0
            # pb描述
            self._pb_desc = pb_desc
            # field描述
            self._field_desc = None
            # 当前序号
            self._order_number = 0
            # 对应的列号
            self._col_array = []

    def __str__(self):
        node_str = self._name
        if self._type != "":
            node_str += "[{}]".format(self._type)
        if self._order_number > 0:
            node_str += "{{order={}}}".format(self._order_number)
        if self._repeated_count > 1 or self._is_repeated:
            node_str += "{repeated="  + str(self._repeated_count) + "}"
        if len(self._col_array):
            node_str += "{{{}}}".format(' '.join([decimal2az(col) for col in self._col_array]))
        return node_str

    def AppendNode(self, str_node_name, str_type_name, str_comment, col):
            """增加一个节点。str_node_name可以是父节点到叶节点的路径，'.'隔开。函数会自动追踪到这些父节点"""
            # 不填名字，就忽略该列
            if str_node_name == "":
                return
            # 分割名称
            node_name_list = str_node_name.split('.', 1)
            # 分割类型, 由于类型可以省略所以暂时不pop
            type_name_list = str_type_name.split('|', 1)
            # 查找节点
            node = self.__get_node_by_name(node_name_list[0])
            # 未找到节点，插入新节点
            if node == None:
                if not self._is_common_type and type_name_list[0] == "":
                    raise Exception(node_name_list[0] + u"不是公共类型，但又没有指定类型名称. 列：" + decimal2az(col))
                node = self.__create_subnode(node_name_list[0], type_name_list[0], str_comment, len(node_name_list) == 1)
                type_name_list.pop(0)
            else:
                # 和当前类型匹配，删除类型名
                real_type_name, _, _ = self.__split_type(type_name_list[0])
                if node._type == real_type_name:
                    type_name_list.pop(0)
                # 结构体是否是repeat只看第一个字段是否重复
                if len(node_name_list) == 1 and node._parent._sub_nodes[0] == node:
                    subnode = node
                    while subnode._parent:
                        if not subnode._not_repeated and subnode._parent._sub_nodes[-1] == subnode:
                            break
                        subnode = subnode._parent
                    subnode._repeated_count += 1
            # 非叶节点，则继续递归遍历
            if len(node_name_list) > 1:
                node.AppendNode(node_name_list[1], '|'.join(type_name_list), str_comment, col)
            else:
                # 如果到了叶节点，但是类型还没用完，表明中间出错了
                if len(type_name_list) != 0 and type_name_list[0] != "":
                    raise Exception("第{0}列，{1}是叶节点，但类型{2}还有剩余。".format(decimal2az(col), node._name, '.'.join(type_name_list)))
                node._col_array.append(col)

    def CheckAndFinish(self):
        if self.__is_message():
            max_order = max([sub_node._order_number for sub_node in self._sub_nodes]) + 1
            for node in self._sub_nodes:
                # 填子类型中的序列号
                if node._order_number == 0:
                    node._order_number = max_order
                    max_order += 1
                if node._repeated_count > 1:
                    node._is_repeated = True
                if node.__is_message():
                    node.CheckAndFinish()

    def Dump(self, tab_num = 0):
        print("    " * tab_num + str(self))
        tab_num += 1
        for node in self._sub_nodes:
            node.Dump(tab_num)

    def GenProto(self, pb_file_name):
        # 输出proto
        output = "/**\n"
        output += "* @file:   " + pb_file_name + "\n"
        output += "* @author: ryanbai(bairuizhi@foxmail.com)\n"
        output += "* @brief:  这个文件是通过工具自动生成的，不要手动修改\n"
        output += "*/\n\n"
        output += DescTree.gen_comm_module() + "\n"
        output += "package op.pb;\n\n"
        output += self.__gen_node_desc()
        output += "\nmessage {self._type}_ARRAY {{\n    repeated {self._type} items = 1;\n}}".format(**vars())

        # 写pb文件
        pb_full_path = PROTO_GEN_PATH + pb_file_name
        pb_file = open(pb_full_path, 'wb+')
        pb_file.write(output)
        pb_file.close()

    def RecheckFieldDesc(self, pb_desc = None):
        if pb_desc:
            self._pb_desc = pb_desc
        for node in self._sub_nodes:
            node._field_desc = self._pb_desc.fields_by_name[node._name]
            # 不是公共类型，并且是message，重新赋值
            if node.__is_message() and not node._is_common_type:
                node._pb_desc = self._pb_desc.nested_types_by_name[node._type]
                node.RecheckFieldDesc()

    def ParseData(self, item, row_values, repeated_num = 0):
        # 遍历子节点
        for node in self._sub_nodes:
            # 数组
            if node._is_repeated:
                for num in range(repeated_num * node._repeated_count, repeated_num * node._repeated_count + node._repeated_count):
                    # struct结构的数组
                    if node.__is_message():
                            # 默认值
                            if not node._nokey and node.__first_child_is_default_value(row_values, num):
                                #print("读到默认值，忽略该结构，继续循环")
                                continue
                            struct_item = item.__getattribute__(node._name).add()
                            node.ParseData(struct_item, row_values, num)
                            if node._nokey and node.__is_all_default(struct_item):
                                item.__getattribute__(node._name).__delitem__(-1)
                                #print("该结构全部是默认值，忽略，继续循环")
                                continue
                    else:
                        cell_value = node.__get_node_value(row_values, num)
                        if cell_value == "":
                            continue
                        if node._field_desc.cpp_type in [1, 2, 3, 4, 7, 8]:
                            # 单列分割
                            if node._repeated_count == 1:
                                item.__getattribute__(node._name).extend([int(float(x)) for x in str(cell_value).split(';')])
                            else:
                                item.__getattribute__(node._name).append(int(float(cell_value)))
                        elif node._field_desc.cpp_type in [5, 6]:
                            # 单列分割
                            if node._repeated_count == 1:
                                item.__getattribute__(node._name).extend([float(x) for x in cell_value.split(';')])
                            else:
                                item.__getattribute__(node._name).append(float(cell_value))
                        else:
                            item.__getattribute__(node._name).append(cell_value)
            # 结构体
            elif node.__is_message():
                struct_item = item.__getattribute__(node._name)
                node.ParseData(struct_item, row_values, repeated_num)
            else:
                cell_value = node.__get_node_value(row_values, repeated_num)
                # 没有值，忽略
                if cell_value == "":
                    continue
                # int类型
                if node._field_desc.cpp_type in [1, 2, 3, 4, 7, 8]:
                    item.__setattr__(node._name, int(cell_value))
                # float类型
                elif node._field_desc.cpp_type in [5, 6]:
                    item.__setattr__(node._name, float(cell_value))
                # 字符串类型
                else:
                    item.__setattr__(node._name, cell_value)

    @staticmethod
    def gen_comm_module():
        return '\n'.join([ "import \"" + common_module + "\";" for  common_module in DescTree.__used_modules])

    @staticmethod
    def __is_type_defined(type_name):
        """该类型是否已经定义过"""
        # 空，直接返回
        if type_name == "":
            return False, None
        # 遍历公共模块
        for module in comm_loaded_modules:
            if type_name in module.DESCRIPTOR.message_types_by_name:
                DescTree.__used_modules.add(module.DESCRIPTOR.name)
                return True, module.DESCRIPTOR.message_types_by_name[type_name]
        return False, None

    def __is_message(self):
        return len(self._sub_nodes) > 0

    def __split_type(self, type_name):
        real_type_name = ""
        default_val = None
        features = ""
        import re
        match_result = re.split(r'[\[\]\=]', type_name)
        if len(match_result) == 2:
            real_type_name = match_result[0]
            default_val = match_result[1]
        elif len(match_result) == 3:
            real_type_name = match_result[0]
            features = match_result[1]
        elif len(match_result) == 4:
            real_type_name = match_result[0]
            features = match_result[1]
            default_val = match_result[3]
        else:
            real_type_name = match_result[0]
        return real_type_name, default_val, features

    def __init_private_type(self, type_name):
        """初始化类型"""
        self._default = None
        self._type = ""
        self._is_repeated = False
        self._is_date = False
        self._is_hour = False
        self._nokey = False
        self._not_repeated = False
        self._type, self._default, features = self.__split_type(type_name)
        # 设置特性
        for feature in [x.strip() for x in features.split(',')]:
            if feature == "repeated":
                self._is_repeated = True
            elif feature == "DateTime":
                self._is_date = True
            elif feature == "HourTime":
                self._is_hour = True
            elif feature == "nokey":
                self._nokey = True
            elif feature == "norepeated":
                self._not_repeated = True

    def __get_node_by_name(self, node_name):
        """根据名称获取子节点"""
        node_list = [node for node in self._sub_nodes if node._name == node_name]
        return node_list[0] if len(node_list)==1 else  None

    def __create_subnode(self, curr_node_name, curr_type_name, str_comment, is_leaf):
        node = DescTree(curr_node_name, curr_type_name, str_comment)
        node._parent = self
        self._sub_nodes.append(node)
        # 继承自父节点，是否是公共类型
        node._is_common_type = self._is_common_type
        # 确定是否是公共类型
        if not node._is_common_type and node._type != "":
            node._is_common_type, node._pb_desc = DescTree.__is_type_defined(node._type)
        elif self._is_common_type and not is_leaf:
            type_name = self._pb_desc.fields_by_name[node._name].message_type.name
            _, node._pb_desc = DescTree.__is_type_defined(type_name)

        # 当前非叶节点不是公共类型，必然是内部类型 
        if not is_leaf and not node._is_common_type and self._pb_desc:
            if node._type in self._pb_desc.nested_types_by_name:
                node._pb_desc = self._pb_desc.nested_types_by_name[node._type]
    
        # 节点的域描述
        if self._pb_desc:
            if node._name in self._pb_desc.fields_by_name:
                node._field_desc = self._pb_desc.fields_by_name[node._name]
                node._order_number = node._field_desc.number
                # 公共类型，且是repeated的，则强制指定
                if self._is_common_type and node._field_desc.label == 3:
                    node._is_repeated = True
            elif self._is_common_type:
                raise Exception("subnode %s is not one field of %s." % (node._name, self._type))

        return node

    def __gen_node_desc(self, tab_num = 0, local_defined_struct = {}):
        """生成pb文件"""
        output = ""
        # 已经定义过
        if self._type in local_defined_struct:
            return output
        # message类型，遍历其子节点
        if self.__is_message()  and not self._is_common_type:
            output = "    " * tab_num + "message " + self._type + " {\n"
            for node in self._sub_nodes:
                # 是消息类型，生成内部类型
                if node.__is_message() and not self._is_common_type:
                    output += node.__gen_node_desc(tab_num + 1, local_defined_struct)
            for node in self._sub_nodes:
                constrait_name = "repeated" if node._is_repeated or node._repeated_count > 1 else "optional"
                default_value = "" if not node._default else "[default = " + str(node._default) + "]"
                if not node.__is_message() and node._comment != "":
                    output += " " * 4 * (tab_num + 1) + "/***\n"
                    output += " " * 4 * (tab_num + 1) + node._comment.replace('\n', '\n' + ' ' * 4 * (tab_num + 1)) + '\n'
                    output += " " * 4 * (tab_num + 1) + "***/\n"
                output += "    " * (tab_num + 1) + "{constrait_name} {node._type} {node._name} = {node._order_number} {default_value};\n".format(**vars()) 
            output += "    " * tab_num + "}\n"
            local_defined_struct[self._type] = 1
        return output

    def __first_child_is_default_value(self, row_values, num):
        # 一直遍历首个子节点
        if self._sub_nodes:
            return self._sub_nodes[0].__first_child_is_default_value(row_values, num)
        # 叶节点是repeated的，不做默认值判断
        if self._is_repeated:
            return False
        # 获取到单元格内容
        cell_value = self.__get_node_value(row_values, num)
        if str(cell_value) == "":
            return True
        return self._field_desc.default_value == cell_value

    def __is_all_default(self, item):
        """判断是否全部为默认值"""
        for (desc, val) in item._fields.items():
            # 数组,长度不为0则不是默认值
            if desc.label == 3:
                if val and len(val) != 0:
                    return False
            elif desc.message_type:
                if not self.__is_all_default(val):
                    return False
            elif desc.default_value != val:
                return False
        return True

    def __get_node_value(self, row_values, num):
        """获取单元格内容"""
        cell_value = row_values[self._col_array[num]]
        #print("列单元格(%s)值为%s" % (decimal2az(self._col_array[num]), str(cell_value)))
        if str(cell_value).strip() == "":
            return ""
        elif self._is_repeated and self._repeated_count == 1:
            return cell_value
        try:
            # 整数
            if self._field_desc.cpp_type in [1, 2, 3, 4, 7, 8]:
                try:
                    tmp = int(float(cell_value))
                    if tmp == 0:
                        return 0
                except:
                    pass 
                # 支持时间类型
                if self._is_date:
                    import time
                    time_struct = time.strptime(cell_value, "%Y-%m-%d %H:%M:%S")
                    return int(time.mktime(time_struct))
                elif self._is_hour:
                    import time
                    time_base = "2000-01-01 "
                    time_struct = time.strptime(time_base + cell_value, "%Y-%m-%d %H:%M:%S")
                    time_struct_base = time.strptime(time_base + "00:00:00", "%Y-%m-%d %H:%M:%S")
                    return int(time.mktime(time_struct)) - int(time.mktime(time_struct_base))
                return int(float(cell_value))
            # 浮点数
            elif self._field_desc.cpp_type in [5, 6]:
                return float(cell_value)
        except:
            raise BaseException("列单元格({})内容错误({})，请检查".format(decimal2az(self._col_array[num]), cell_value))
        # string或bytes
        if str(cell_value).endswith('.0'):
            try:
                tmp = int(float(cell_value))
                return unicode(tmp)
            except:
                pass
        return unicode(cell_value)

# protoc 
PROTOC_PATH = PROTOC + " -I" + PB_PROTO + " -I" + CUSTOM_PROTO
# 输出文件前缀
OUTPUT_FILE_BASE="dataconfig_"
# proto路径
create_path_if_noexist(PROTO_GEN_PATH)
# python导出路径
create_path_if_noexist(PYTHON_GEN_PATH)
# 数据文件
create_path_if_noexist(DATA_GEN_PATH)
# 文本文件
create_path_if_noexist(TEXT_GEN_PATH)

# 前几行固定用途
FIELD_TYPE_ROW = 0
FIELD_NAME_ROW = 1
FIELD_COMMENT_ROW = 2

class SheetInterpreter:
    """通过excel配置生成配置的protobuf定义文件"""

    def __init__(self, xls_file, sheet_name):
        """指定excel表和页签列表"""
        self._sheet_type_name = sheet_name
        # 打开所有页签
        workbook = xlrd.open_workbook(xls_file)
        self._sheet = workbook.sheet_by_name(sheet_name)
        # proto输出
        self._pb_file_name = OUTPUT_FILE_BASE + self._sheet_type_name.lower() + ".proto"
        # py输出
        self._py_module_name = OUTPUT_FILE_BASE + self._sheet_type_name.lower() + "_pb2"
        # data
        self._data_file_name = OUTPUT_FILE_BASE + self._sheet_type_name.lower() + ".data"
        # txt
        self._txt_file_name = OUTPUT_FILE_BASE + self._sheet_type_name.lower() + ".txt"
        #
        self.module = None
        # 生成python格式
        self.__export_proto()
        # 描述树
        self._desc_tree = DescTree(self._sheet_type_name.lower(), self._sheet_type_name, self._sheet_type_name, self.__find_desc_of_exist_pb())
        # 
        self._begin_row = FIELD_COMMENT_ROW + 1

    def Interpreter(self) :
        """生成proto和数据"""
        #通过第一个页签导出
        type_sheet = self._sheet
        # 行数太少
        if type_sheet.nrows <= FIELD_COMMENT_ROW:
            raise Exception("{type_sheet.name}中只有{type_sheet.nrows}行，不符合格式要求".format(**vars()))

        #print("开始导出%s, excel表共%d列, 开始行:%u" % (self._pb_file_name, type_sheet.ncols, self._begin_row))
        for col in range(type_sheet.ncols):
            node_name = type_sheet.cell_value(FIELD_NAME_ROW, col).strip()
            node_type = type_sheet.cell_value(FIELD_TYPE_ROW, col).strip()
            node_desc = type_sheet.cell_value(FIELD_COMMENT_ROW, col).strip()
            #print("第%s列， 名称：%s 类型：%s 注释：%s" % (decimal2az(col), node_name, node_type, node_desc))
            self._desc_tree.AppendNode(node_name, node_type, node_desc, col)
        self._desc_tree.CheckAndFinish()
        self._desc_tree.Dump()

        #  导出proto
        self._desc_tree.GenProto(self._pb_file_name)

        # 重新生成proto的python描述
        self.__export_proto(True)
        self._desc_tree.RecheckFieldDesc(self.__find_desc_of_exist_pb())

        # 导出数据
        # 找到array类型
        item_array = getattr(self.module, self._sheet_type_name+'_ARRAY')()
        print("开始导出页签%s中的数据" % self._sheet.name)
        for row in range(self._begin_row, self._sheet.nrows):
            print("开始导出%s第%u行" % (self._sheet.name, row))
            self._desc_tree.ParseData(item_array.items.add(), self._sheet.row_values(row))

            # 写data文件
            data_file = open(DATA_GEN_PATH + self._data_file_name, 'wb+')
            data_file.write(item_array.SerializeToString())
            data_file.close()

            # 写text文件
            text_file = open(TEXT_GEN_PATH + self._txt_file_name, 'wb+')
            from google.protobuf.text_format import MessageToString
            text_file.write(MessageToString(item_array, True)) 
            text_file.close()

    def __export_proto(self, must_succ = False):
        # 生成python和cpp
        full_pb_file = PROTO_GEN_PATH + self._pb_file_name
        if os.path.exists(full_pb_file):
            command = PROTOC_PATH + " -I{0} --python_out={1} ".format(PROTO_GEN_PATH, PYTHON_GEN_PATH) + full_pb_file
            #print(command)
            os.system(command)
        elif must_succ:
            raise Exception("文件{0}不存在.".format(full_pb_file))

    def __find_desc_of_exist_pb(self):
        # python文件也放svn，防止错误
        pb_py_full_path = PYTHON_GEN_PATH + self._py_module_name + ".py"
        if os.path.isfile(pb_py_full_path):
            # remove loaded module
            if self.module:
                os.system("rm " + pb_py_full_path + "c")
                del sys.modules[self._py_module_name]
            self.module = __import__(self._py_module_name)
            return self.module.DESCRIPTOR.message_types_by_name[self._sheet_type_name]
        return None

if __name__ == '__main__' :
    """入口"""
    if len(sys.argv) < 3 :
        print("Usage: %s sheet_name|sheet_name(should be upper) xls_file" % sys.argv[0])
        sys.exit(-1)

    try:
        parser = SheetInterpreter(sys.argv[2], sys.argv[1])
        parser.Interpreter()
    except:
        traceback.print_exc()
        sys.exit(-1)
    else:
        sys.exit(0)

