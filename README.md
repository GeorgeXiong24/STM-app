# Smart-Testing-of-Memorization-App

For most English-Chinese word reciting or remembering app in the market currently, the word dictionary or word list is usually being limited or already being set up including a specific range. 

This app is focusing on giving the capability of customization to users, where they are able to upload their own word list in .xlsx or .numbers format onto our app locally, and the words with their Chinese definitions can be recognized automatically, no matter the arrangement or format.

Moreover, the Chinese definition of an English word may vary, almost never limited to a single one. Consequently, the way our app test how the user master those English words cannot simply comparing users' input and the correct def demonstrated within the uploaded list. To solve this problem, we provoke the API of DeepSeek in order to help distinguishing whether the input definition is similar to the correct one, which provides a better experience for users.

After a full testing, a report will be generated, demonstrating the number of times the users get each word incorrect or give up trying, which gives them an opportunity to get to know their current ability of mastering those words.

Finally, you have two options to use this app. One is to buy the DeepSeek API yourself, and enter the API into the input box at the start of the app interface. This allows the app to directly provoke the AI to function normally. The other option is currently under development, if it is finished, we will update the .md file to inform you about the improvement. Other functions are also under development, please wait for more functions to come out.

Thank you so much for using and supporting our app.
