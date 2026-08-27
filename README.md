# Word-Reciting-App

For most English-Chinese word reciting or remembering app in the market currently, the word dictionary or word list is usually being limited or already being set up. 

This app is focusing on giving the capability of customization to the users, where they are able to upload their own word list in .xlsx or .numbers format onto our app locally, and the words with their Chinese definitions can be recognized. Since the upload of the file is locally, so there is no worry for data leakage. 

Currently, the recognition of word list is only available for a particular format of arrangement of the list, where one column named "单词" will be concluded as words, and the column named "解释" will be concluded as Chinese definitions. We are currently improving on this aspect, allowing smarter analysis.

Moreover, the Chinese definition of an English word may vary, almost never limited to a single one. Consequently, the way our app test how the user master those English words cannot simply comparing users' input and the correct def demonstrated within the uploaded list. To solve this problem, we provoke the API of DeepSeek in order to help distinguishing whether the input definition is similar to the correct one, which provides a better experience for users.

Finally, you have two options to use this app. One is to buy the DeepSeek API yourself, and enter the API into the input box at the start of the app interface. This allows the app to directly provoke the AI to function normally. The other option is currently under development, if it is finished, we will update the .md file to inform you about the improvement.

Thank you so much for using and supporting our app.
