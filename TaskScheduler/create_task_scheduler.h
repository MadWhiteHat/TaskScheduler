#ifndef _CREATE_TASK_SCHEDULER_H
#define	_CREATE_TASK_SCHEDULER_H

#include <iostream>
#include <vector>

#include <windows.h>
#include <taskschd.h>

struct ITask {
public:
    struct EventData {
        LPCWSTR _name = nullptr;
        LPCWSTR _queryString = nullptr;
    };

    static BOOL SchedulerInit(ITaskService** __task);
    static void SchedulerDeinit(ITaskService** __task);

    static BOOL CreateTask(
        ITaskService** __service,
        ITaskDefinition** __task,
        LPCWSTR __authorName
    );
    static BOOL AddTriggerEvent(
        ITaskService** __service,
        ITaskDefinition** __task,
        LPCWSTR __triggerDelay,
        LPCWSTR __triggerName,
        LPCWSTR __triggerQueryString,
        std::vector<ITask::EventData>& __triggerValueQueries
    );
    static BOOL AddExecAction(
        ITaskDefinition** __task,
        LPCWSTR __taskExecName,
        std::vector<ITask::EventData>& __taskValueQueries
    );
    static BOOL RegisterTask(
        ITaskService** __service,
        ITaskDefinition** __task,
        LPCWSTR __folder,
        LPCWSTR __taskName
    );

    static void IRelease(IUnknown** __iface);

protected:

    static void _PrintError(const char* __msg, DWORD __errorCode);

private:
};

#endif // !_CREATE_TASK_SCHEDULER_H

