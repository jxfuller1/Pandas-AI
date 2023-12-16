I created this program to use AI to parse data out of excel files using prompts.

Created from pandasai version 1.5.5.... the versions are changing constantly at the moment.  it utilizes the
Agent feature in pandasai and Googles Palm2 AI (which is free).  Just add your own API key to it.

![image](https://github.com/jxfuller1/Pandas-AI/assets/123666150/7116f792-89b2-4b53-bd87-52fa52b13be4)


Take note that the AI calls and reading the excel to a dataframe are offloaded to a QThread to a keep the UI responsive
as the AI does it's thing.  Therefore, I have to pass the pandasai agent back and forth between the main thread and Qthread.
I put a number of try statements in, especially around the AI function calls for any issues that may arise.
