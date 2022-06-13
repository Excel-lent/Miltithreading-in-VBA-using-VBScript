# Multithreading in VBA using VBScript.
This repository is based on an article "[Multithreading VBA using VBscript](https://analystcave.com/excel-multithreading-in-vba-using-vbscript/)". I tried to do multithreading VBA more clear and structured. The main idea of the method is pretty simple: the VBScript can be executed from VBA and the communication between started VBScript and VBA is accomplished via Excel cells. The communication between VBScript and VBA is necessary because we need to know when the VBScript task is finished and to obtain the results of its execution.

The code consist of three parts:
1. The main VBA module with two functions:
    * **RunAllTasksAsynchronously**, that initializes and starts the threads and afterwards starts **MainLoop**. Each thread in example is initialized with execution time to show that the different threads will be started and finished at different times. So, execution time of "Thread 1" is set to 1 second, execution time of "Thread 2" is set to 2 seconds ... and execution time of "Thread N" is set to 2^N seconds. 
    * **MainLoop** - checks the state of all threads in infinite loop and if the thread is finished, do some work with execution results (in example the results are copied to other cell) and starts the thread again. Execution is finished when all tasks are finished.
2. Class "Thread" contains:
    * **Constructor** that simply copied some never changed variables (execution time, workbook and worksheet names) to global variables of the class.
    * **StartVBScriptThread** - constructs arguments for VBScript and runs it in background.
    * **GetThreadState** - returns current state of the thread.
3. VBScript ("Thread.vbs") takes input arguments and waits predefined number of seconds. The input arguments are:
    * **Workbook name** - "MultithreadingViaVBS v0.02.xlsm".
    * **Worksheet name** - "Threads".
    * **ThreadID** - Number from 1 to maximal number of threads.
    * **Input parameter 1** - execution time in seconds.
    * **Output cell** - corresponding cell in column "F". The output of VBScript is "Greetings from thread N".
    * **Thread state cell** - corresponding cell in column "E".