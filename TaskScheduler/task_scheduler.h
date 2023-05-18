#ifndef _TASK_SCHEDULER_H
#define _TASK_SCHEDULER_H

#include <stack>

#include <taskschd.h>

#define OUTPUT_LEVEL '\t'

void PrintTasks(void);
BOOL SchedulerInit(ITaskService** __service);
void SchedulerDeinit(void);
void ShowFolderTask(std::stack<ITaskFolder*>& __folders);
void ShowTasksInfo(
    IRegisteredTaskCollection* __taskCollection,
    LONG __tasksCount
);
BOOL PushSubFolders(
    std::stack<ITaskFolder*>& __folders,
    ITaskFolderCollection* __taskFoldersCollection
);

void PrintError(const char* __msg, DWORD __errorCode);

#endif // _TASK_SCHEDULER_H