# STM — Smart Testing of Memorization

> An AI-powered memorization tester that works with **your own** word lists.

[English](README.md) · [中文](README_zh.md)

## What is STM?

STM is a desktop app that tests how well you have memorized a list of vocabulary words. Instead of a fixed dictionary or multiple-choice questions, STM lets you bring your own word list and answer in your own words.

- Upload your own `.xlsx` or `.numbers` word list.
- AI automatically recognizes each word and its Chinese definition, no matter how the list is arranged.
- Because a word can have more than one valid meaning, AI judges whether your answer matches the meaning — not just the exact text.
- After testing, you get a report of the words you missed or gave up on, which you can export for review.

## Supported languages

Words can be in any of these languages:

Chinese · French · German · Spanish · Italian · Portuguese · Japanese · Korean

Definitions are shown and tested in **Chinese**.

## How it works

1. **Recognition** — When you upload a spreadsheet, STM asks the DeepSeek AI to identify each word and its Chinese definition, regardless of the column names or layout.
2. **Judging** — During the test you type the definition in Chinese. The AI compares your answer to the reference meaning, accepting synonyms, natural paraphrases, and minor typos.
3. **Review** — Words answered incorrectly are retested once more at the end.

## Getting started

1. Get a DeepSeek API key (it starts with `sk-`).
2. Download the latest version from the [Releases](https://github.com/GeorgeXiong24/STM-app/releases) page.
3. Open STM and paste your API key, then click **OK**.
4. Choose the language of your words.
5. Drop a `.xlsx` or `.numbers` file into the window, or use **File → Open…**.
6. Review the recognized words and delete any rows you do not want.
7. Click **Go** to start the test.
8. For each word, type its Chinese definition and press **Enter** or click **OK**.
9. When finished, review the report and export any words you missed.

## The test

- Each word gets two attempts.
- Click **Give up** to skip a word you do not know.
- Correct answers are filled in automatically.
- Missed words are shown again in a final review round.

## Exporting results

After a test, export the incorrect or given-up words as `.xlsx` or `.numbers`, with an optional column showing how many times each word was missed.

## Notes

- Uploaded files are stored only temporarily and are removed when you close STM.
- Press **Esc** to return to the previous screen.
- On macOS, see the [Mac helper](https://github.com/GeorgeXiong24/STM-app/blob/main/Mac_helper.md) if you run into any issues.

## Feedback

Found a bug or have a suggestion? Open an issue in the [issue tracker](https://github.com/GeorgeXiong24/STM-app/issues).

## Download

Get the latest version from the [Releases](https://github.com/GeorgeXiong24/STM-app/releases) page.
