#! /usr/bin/env python
#coding=utf-8
PROTOC="/usr/local/protobuf-2.6.1/src/protoc"
PB_PROTO="/usr/local/protobuf-2.6.1/src"
PB_PYTHON="/usr/local/protobuf-2.6.1/python"
#自己项目所产生的proto
CUSTOM_PYTHON="./protobuf/python"
CUSTOM_PROTO="./protobuf"
PROTO_GEN_PATH="./proto/"
# python导出路径
PYTHON_GEN_PATH="./python/"
# 数据文件
DATA_GEN_PATH="./deploy_data/"
# 文本文件
TEXT_GEN_PATH="./readable_data/"
# 依赖公共模块conf_struct.proto
COMMON_MODULES=['conf_struct_pb2']
