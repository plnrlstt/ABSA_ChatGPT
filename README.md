# ChatGPT as a PR-Management Expert

This is the repository for the group project of the lecture `Formal Semantics 2023/24` at the Computational Linguistics Faculty of Heidelberg University. 

Our project is about using **ChatGPT as a PR-Management Expert** that evaluates restaurant reviews by creating sentiment quads, writes weekly reports and issues warnings for critical situations.

In our [project outline](https://docs.google.com/document/d/12RYexcdQH-3xhM-T3H6ztGkI95sITbloao5V_5Vknhw/edit) we summarized our inital thoughts and ideas and also how we planned to approach the task. The [final project report](Group_BEP__Report.pdf) gives a detailed insight into the execution of out project,our results and the conclusions. 

The explanation of the code and how to use it can be found [here](code/Group_BEP__Code.pdf).

## 🧾Contents

1. [About](#introduction)
2. [Requirements](#paragraph1)
5. [Group BEP members](#paragraph4)

## 🤖About <a name=introduction></a>

The aim of our project is to test ChatGPT’s abilities to provide an insightful analysis of customer feedback as a tool for management purposes, which includes an ASQP task, generation of weekly summaries of given reviews, and production of warnings in case of abnormalities in the dataset. For this purpose, we are utilizing an annotated [dataset](https://github.com/IsakZhang/ABSA-QUAD/tree/master), which serves as our gold standard, enabling us to compare ChatGPT's output against it.

## 💻Requirements <a name=paragraph1></a>

For the Python code used in this repository you need to fulfill some basic requirements.

* Visual Studio Code (or any other programming environment)
* OpenAI API Key (can be obtained [here](openai.com))
* Excel

You also need to install some libraries which can be performed with the following line:

```python
pip install pandas openai openpyxl xlsxwriter
```

Additional to the code, you also need to download the .env file in this folder.

## 👩‍👩‍👦Group BEP members <a name=paragraph4></a>
* Lukas Dorschner (lukasdorschner@yahoo.de)
* Polina Kuznetcova (polina.kuznetcova@stud.uni-heidelberg.de)
* Begüm Yildiz (begumyildiz@web.de)
