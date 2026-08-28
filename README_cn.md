# Smart-Testing-of-Memorization-App

英文版: https://github.com/GeorgeXiong24/STM-app/blob/main/README.md
中文版: https://github.com/GeorgeXiong24/STM-app/blob/main/README_cn.md

目前市场上大多数英汉单词背诵应用都采用固定或限定的单词词典或列表，而且形式基本上都是选择题，这就导致蒙对的几率比较高，从而降低了真正的掌握能力和效率。该应用则给予了用户进行自定义的权利：用户可以将自己的单词列表以 .xlsx 或 .numbers 格式上传到应用中。无论单词的排列或格式如何，AI都能自动识别这些单词及其对应的中文释义。

此外，英语单词的中文释义可能有所不同，通常不会局限于单一释义。因此，该应用并不会仅通过比较用户输入的内容与上传文件中的正确释义是否相同来判断用户是否真正掌握了这些单词。为了解决这个问题，我们使用了DeepSeek的API来判断用户输入的释义是否与正确释义相似，从而为用户提供更好的使用体验。

经过完整的测试后，该应用会生成一份报告，显示用户每个单词错误或放弃尝试的次数。这样用户可以了解自己当前掌握这些单词的情况。此外，用户还可以导出包含错误使用或放弃尝试的单词、其中文释义的 .xlsx 或 .numbers 文件，并决定是否显示这些次数。

最后，要使用此应用，您需要自行购买DeepSeek API，然后将其输入到应用开始界面的输入框中，这样应用就能直接让AI正常运作。更高级的功能正在开发中，请耐心等待。如果有任何问题、漏洞以及改进建议，欢迎在 https://github.com/GeorgeXiong24/STM-app/issues 中提交。

如需下载该应用，请访问 https://github.com/GeorgeXiong24/STM-app/releases 。非常感谢您使用和支持我们的应用。
