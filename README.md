<img width="305" height="193" alt="image" src="https://github.com/user-attachments/assets/bf7ee433-fd1c-4a1f-910d-c43acadaf808" />

Press Start button, the tools will try to launch Excel file c:\test\sample.xlsx,
please put the excel you want to launch under c:\test with file name sample.xlsx,
you could adjust the interation number to launch the excel couple times,
the tools will auto close the excel after launch and average the launch time,


<img width="596" height="354" alt="image" src="https://github.com/user-attachments/assets/07942d4d-986c-447b-9951-eda01e38338a" />

the launch time consist of two parts
-file open, launch excel + load file + load macro
-rendering time, extra time about rendering/update some graph/UI ex: excel working on suggeting different chart types even the raw data of chart had been load
you would see pretty long rendering time, if the excel file contain charts


the Camera button is not related to excel launch time but active the camera without launch any camera application
(which is used for another appliction)

excutable could be found under bin folder


