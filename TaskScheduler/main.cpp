#include <iostream>
#include "task_scheduler.h"
#include "create_task_scheduler.h"

void Usage() {
    std::cout << "1. Print tasks" << std::endl;
    std::cout << "2. Create Firewall Task" << std::endl;
    std::cout << "3. Create ping task" << std::endl;
    std::cout << "4. Exit" << std::endl;
}

void CreateFirewallTask(LPCWSTR __taskName) {
    ITaskService* __service = nullptr;
    ITaskDefinition* __task = nullptr;
    BOOL __res = FALSE;

    __res = ITask::SchedulerInit(&__service);
    if (!__res) { return; }

    __res =  ITask::CreateTask(&__service, &__task, L"Kozelkov Mike");
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }

    std::vector<ITask::EventData> __queryValues;
    __queryValues.push_back(
        ITask::EventData{
            L"event_data",
            L"Event/System/EventID"
        }
    );

    CONST WCHAR __queryAdvSec[] =
        L"<QueryList>"
        L"<Query Id='0'>"
        L"<Select Path='Microsoft-Windows-Windows Firewall With Advanced Security/Firewall'>"
        L"*[System[Provider[@Name='Microsoft-Windows-Windows Firewall With Advanced Security'] and EventID=2003]]"
        L"</Select>"
        L"</Query>"
        L"</QueryList>";


    __res = ITask::AddTriggerEvent(
        &__service,
        &__task,
        L"P0Y0M0DT0H0M10S",
        L"Advanced Security",
        __queryAdvSec,
        __queryValues
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }

    CONST WCHAR __queryWHC[] =
        L"<QueryList>"
        L"<Query Id='0'>"
        L"<Select Path='Microsoft-Windows-Windows Defender/WHC'>"
        L"*[System[Provider[@Name='Microsoft-Windows-Windows Defender'] and EventID=5007]]"
        L"</Select>"
        L"</Query>"
        L"</QueryList>";


    __res = ITask::AddTriggerEvent(
        &__service,
        &__task,
        L"P0Y0M0DT0H0M10S",
        L"WHC",
        __queryWHC,
        __queryValues
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }

    CONST WCHAR __queryOperational[] =
        L"<QueryList>"
        L"<Query Id='0'>"
        L"<Select Path='Microsoft-Windows-Windows Defender/Operational'>"
        L"*[System[Provider[@Name='Microsoft-Windows-Windows Defender'] and EventID=5007]]"
        L"</Select>"
        L"</Query>"
        L"</QueryList>";


    __res = ITask::AddTriggerEvent(
        &__service,
        &__task,
        L"P0Y0M0DT0H0M05S",
        L"Operational",
        __queryOperational,
        __queryValues
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }

    __res = ITask::AddExecAction(
        &__task,
        L"\\System32\\StatusMessageBox.exe",
        __queryValues
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }

    
    __res = ITask::RegisterTask(
        &__service,
        &__task,
        L"\\",
        __taskName
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }
    ITask::SchedulerDeinit(&__service);
    return;
}

void CreatePingTask(LPCWSTR __taskName) {
    ITaskService* __service = nullptr;
    ITaskDefinition* __task = nullptr;
    BOOL __res = FALSE;

    __res = ITask::SchedulerInit(&__service);
    if (!__res) { return; }

    __res =  ITask::CreateTask(&__service, &__task, L"Kozelkov Mike");
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }

    std::vector<ITask::EventData> __queryValues;
    __queryValues.push_back(
        ITask::EventData{
            L"event_id",
            L"Event/System/EventID"
        }
    );
    __queryValues.push_back(
        ITask::EventData{
            L"event_ip",
            L"Event/EventData/Data[@Name='SourceAddress']"
        }
    );

    CONST WCHAR __queryPing[] =
        L"<QueryList><Query Id='0' Path='Security'>"
        L"<Select Path='Security'>"
        L"*[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and EventID=5152]]"
        L"</Select>"
        L"</Query>"
        L"</QueryList>";


    __res = ITask::AddTriggerEvent(
        &__service,
        &__task,
        L"P0Y0M0DT0H0M0S",
        L"Ping",
        __queryPing,
        __queryValues
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }
    __res = ITask::AddExecAction(
        &__task,
        L"\\System32\\StatusMessageBox.exe",
        __queryValues
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }

    __res = ITask::RegisterTask(
        &__service,
        &__task,
        L"\\",
        __taskName
    );
    if (!__res) {
        ITask::SchedulerDeinit(&__service);
        return;
    }
    ITask::SchedulerDeinit(&__service);
    return;
}

int main(void) {
    setlocale(LC_ALL, "Russian");
    int choice = 0;
    while (true) {
        system("cls");
        Usage();
        std::cout << "Enter choice: ";
        std::cin >> choice;
        switch (choice) {
            case 1 :
                PrintTasks();
                break;
            case 2 : 
                CreateFirewallTask(L"BSIT3_Firewall_Task");
                break;
            case 3 :
                CreatePingTask(L"BSIT3_Ping_Task");
                break;
            case 4:
                return 0;
            default:
                std::cout << "Invalid option" << std::endl;
        }
        system("pause");
    }
}