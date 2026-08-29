# Smart-Testing-of-Memorization-App

English version: https://github.com/GeorgeXiong24/STM-app/blob/main/README.md

Chinese version: https://github.com/GeorgeXiong24/STM-app/blob/main/README_cn.md

For most English-Chinese word reciting or remembering app in the market currently, the word dictionary or word list is usually being limited or already being set up including a specific range. 

This app is focusing on giving the capability of customization to users, where they are able to upload their own word list in .xlsx or .numbers format onto our app locally, and the words with their Chinese definitions can be recognized automatically by AI, no matter the arrangement or format.

Moreover, the Chinese definition of an English word may vary, almost never limited to a single one. Consequently, the way our app test how the user master those English words cannot simply compare users' input and the correct def demonstrated within the uploaded file. To solve this problem, we provoke the API of DeepSeek in order to help distinguishing whether the input definition is similar to the correct one, which provides a better experience for users.

After a full testing, a report will be generated, demonstrating the number of times the users get each word incorrect or give up trying, which gives users the opportunity to get to know their current ability of mastering those words. Moreover, users are able to export the file containing the English words they got incorrect or given up, their Chinese definitions, and the users can decide whether the number of times the user got each word incorrect or given up will be displayed or not.

Finally, to use this app, you can buy the DeepSeek API yourself, and enter the API into the input box at the start of the app interface. This allows the app to directly provoke the AI to function normally. More advanced functions are currently under development, please wait for them to come out. If there are any issues or bugs or suggestions for improvement, you are welcome to submit them in https://github.com/GeorgeXiong24/STM-app/issues.

For downloading our app, please visit https://github.com/GeorgeXiong24/STM-app/releases. Thank you so much for using and supporting our app.

Additionally, for macbook users, you may encounter some errors. In order to solve them, you can visit https://github.com/GeorgeXiong24/STM-app/blob/main/Mac_helper.md.
