# XMind 一键转换 Excel

## 克隆项目
```
git clone https://github.com/YeKelvin/xmind-to-excel.git
```

## 安装
```
pip install -r requirements.txt -i https://mirrors.aliyun.com/pypi/simple/
```

## XMind 用例格式
详情请参考 [testcase.xmind](https://github.com/YeKelvin/xmind-to-excel/blob/master/testcase.xmind)

## 使用说明
暂时没有提供命令行调用，直接代码运行吧

打开 `transformer.py` 修改 main 函数:

```python
if __name__ == '__main__':
    xmind_file_path = r'xmind文件路径'
    xmind_sheet_name = 'sheet页名称'
    xmind_to_excel_for_tapd(xmind_file_path, xmind_sheet_name)
```
