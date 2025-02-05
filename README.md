大模型嵌入Word教程，让你的办公效率起飞！

通过 VBA 脚本调用 API
如果你熟悉 VBA（Visual Basic for Applications），可以通过编写脚本调用 AI 模型的 API（如 OpenAI API 或 DeepSeek API、硅基流动平台）。
步骤：
1.在 Word 中打开 VBA 编辑器（按 Alt + F11）。
2.编写 VBA 脚本，使用 HTTP 请求调用 AI 模型的 API。
3.将 API 返回的结果插入到 Word 文档中。

1. 获取API key
- deepseek-r1模型API Key获取地址：https://www.deepseek.com/
- qwen2.5模型API Key获取地址：https://bailian.console.aliyun.com/

2. 配置Word
1. 启用开发者工具
   - 新建Word，点击 文件 -> 选项 -> 自定义功能区
   - 勾选"开发者工具"

2. 配置信任中心
   - 点击 信任中心 -> 信任中心设置
   - 选择"启用所有宏"与"信任对VBA......"

3. 添加模块
   - 点击开发者工具，点击Visual Basic
   - 在新窗口中点击插入，选择模块
   - 将代码复制到编辑区中（**注意替换你自己的API key**）
![image](https://github.com/user-attachments/assets/1735fd66-38eb-42a4-88f9-385590d25baf)
4. 自定义功能区
   - 点击 文件 -> 选项 -> 自定义功能区
   - 右键开发工具，点击添加新组
   - 右键新建组，点击重命名
   - 将其命名为"DeepSeek"，选择心仪的图标
   - 点击确定

![1738747154817](https://github.com/user-attachments/assets/6b2caf31-05b6-4217-b6e5-562001fe1460)

![image](https://github.com/user-attachments/assets/7256012a-2a27-419e-918d-a58bcb6de19f)

5. 添加命令
   - 选择DeepSeek（自定义）
   - 在左侧命令中选择"宏"
   - 找到并选中"DeepSeekV3"，点击添加
   - 右键添加的命令，点击重命名
   - 选择开始符号作为图标
   - 重命名为"生成"

![image](https://github.com/user-attachments/assets/a3e3ebfc-d05d-4b52-90c3-535efd9e61d9)
3. 使用方法
1. 选中需要处理的文字
2. 点击"生成"按钮
3. 等待大模型响应

4. 创建模板
1. 将你的文档另存为 Word 模板 (.dotm)：
   - 点击"文件" → "另存为"
   - 选择保存类型为"Word 启用宏的模板 (.dotm)"
   - 保存到 Word 的模板文件夹（通常是 C:\Users\用户名\AppData\Roaming\Microsoft\Word\STARTUP）
   - 这样每次打开 Word 时，宏就会自动可用




