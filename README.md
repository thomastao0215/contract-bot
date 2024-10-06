# Docx file transfer


## 运行环境变量
| Name       | Default Value | Desc        |
| ---------- | ------------- | ----------- |
| PORT       | 11000         | 运行端口号 |

## 安装
```
npm install
```
## 运行
```
npm run start
```

## 模版1 变量
| Name       | Default Value | Desc        |
|  template  |               | 引用的模版        |


## CURL 运行示例
- 执行下面的 curl 生成的 docs, 里面的内容中 `{name}` 位置已经被替换为 `anxing`
```
curl http://localhost:11000/generate/docx?template=template1&name=anxing
```

