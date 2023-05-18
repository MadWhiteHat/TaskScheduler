#include <iostream>
#include <fstream>
#include <stack>

#include <windows.h>
#include <comdef.h>

#include <taskschd.h>

#include "task_scheduler.h"

void PrintTasks(void) {
    HRESULT __hr;

    ITaskService* __service = NULL;

    if (!SchedulerInit(&__service)) { return; }

    ITaskFolder* __rootFolder = nullptr;

    __hr = __service->GetFolder(_bstr_t(), &__rootFolder);

    __service->Release();

    if (FAILED(__hr)) {
        PrintError("ITaskService::GetFolder  failed", __hr);
        SchedulerDeinit();
        return;
    }

    std::stack<ITaskFolder*> __folders;
    __folders.push(__rootFolder);

    ShowFolderTask(__folders);

    SchedulerDeinit();
}

BOOL SchedulerInit(ITaskService** __service) {
    HRESULT __hr;
    
    __hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(__hr)) {
        PrintError("CoInitializeEx failed", __hr);
        return FALSE;
    }

    __hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL
    );

    if (FAILED(__hr)) {
        PrintError("CoInitializeSecurity failed", __hr);
        CoUninitialize();
        return FALSE;
    }

    __hr = CoCreateInstance(
        CLSID_TaskScheduler,
        NULL,
        CLSCTX_INPROC_SERVER,
        IID_ITaskService,
        reinterpret_cast<void**>(__service)
    );

    if (FAILED(__hr)) {
        PrintError("CoCreateInstance failed", __hr);
        CoUninitialize();
        return FALSE;
    }

    __hr = (*__service)->Connect(
        _variant_t(),
        _variant_t(),
        _variant_t(),
        _variant_t()
    );

    if (FAILED(__hr)) {
        PrintError("ITaskService::Connect failed", __hr);
        (*__service)->Release();
        CoUninitialize();
        return FALSE;
    }
    return TRUE;
}

void SchedulerDeinit() { CoUninitialize(); }

void ShowFolderTask(std::stack<ITaskFolder*>& __folders) {
    HRESULT __hr;

    while (!__folders.empty()) {
        ITaskFolder* __folder = __folders.top();
        __folders.pop();

        IRegisteredTaskCollection* __taskCollection = nullptr;
        __hr = __folder->GetTasks(NULL, &__taskCollection);

        if (FAILED(__hr)) {
            PrintError("ITaskFolder::GetTasks failed", __hr);
            __folder->Release();
            return;
        }

        ITaskFolderCollection* __taskFoldersCollection = nullptr;
        __hr = __folder->GetFolders(0, &__taskFoldersCollection);

        if (FAILED(__hr)) {
            PrintError("ITaskFolder::GetFolders failed", __hr);
            __folder->Release();
            return;
        }

        BSTR __pathToFolder;

        __folder->get_Path(&__pathToFolder);
        __folder->Release();

        LONG __tasksCount = 0;
        __taskCollection->get_Count(&__tasksCount);

        std::wcout << reinterpret_cast<wchar_t*>(__pathToFolder) << std::endl;
        std::wcout << OUTPUT_LEVEL << "Number of tasks: "
            << __tasksCount << std::endl;

        SysFreeString(__pathToFolder);

        ShowTasksInfo(__taskCollection, __tasksCount);
        if (!PushSubFolders(__folders, __taskFoldersCollection)) {
            std::cout << OUTPUT_LEVEL << "Failed to push subfolders"
                << std::endl;
        }

        __taskFoldersCollection->Release();
        __taskCollection->Release();
    }
}

void ShowTasksInfo(
    IRegisteredTaskCollection* __taskCollection,
    LONG __tasksCount
) {
    HRESULT __hr;

    for (LONG i = 0; i < __tasksCount; ++i) {
        std::wcout << OUTPUT_LEVEL << OUTPUT_LEVEL;

        IRegisteredTask* __task = nullptr;
        __hr = __taskCollection->get_Item(_variant_t(i + 1), &__task);

        if (SUCCEEDED(__hr)) {
            BSTR __taskName;
            TASK_STATE __taskState;

            __hr = __task->get_Name(&__taskName);
            if (FAILED(__hr)) {
                std::wcout << L"Cannot obtain name of task ¹" << i + 1
                    << std::endl;
                continue;
            }

            std::wcout << reinterpret_cast<wchar_t*>(__taskName) << " : ";

            __hr = __task->get_State(&__taskState);
            if (SUCCEEDED(__hr)) {
                switch (__taskState) {
                    case TASK_STATE_READY:
                        std::wcout << L"READY" << std::endl;
                        break;
                    case TASK_STATE_RUNNING:
                        std::wcout << L"RUNNING" << std::endl;
                        break;
                    case TASK_STATE_QUEUED:
                        std::wcout << L"QUEUED" << std::endl;
                        break;
                    case TASK_STATE_DISABLED:
                        std::wcout << L"DISABLED" << std::endl;
                        break;
                }
            } else { std::cout << "UNKNOWN" << std::endl; }
        } else {
            std::wcout << L"Cannot obtain info about task ¹"
                << i + 1 << std::endl;
        }
    }
}

BOOL PushSubFolders(
    std::stack<ITaskFolder*>& __folders,
    ITaskFolderCollection* __taskFoldersCollection
) {
    HRESULT __hr;
    LONG __foldersCount = 0;

    __hr = __taskFoldersCollection->get_Count(&__foldersCount);
    if (FAILED(__hr)) { return FALSE; }

    for (LONG i = 0; i < __foldersCount; ++i) {
        ITaskFolder* __subFolder;
        __hr = __taskFoldersCollection->get_Item(
            _variant_t(i + 1),
            &__subFolder
        );
        if (FAILED(__hr)) {
            std::wcout << OUTPUT_LEVEL;
            std::wcout << L"Cannot obtain folder ¹" << i + 1 << std::endl;
            continue;
        }
        __folders.push(__subFolder);
    }
    return TRUE;
}

void PrintError(const char* __msg, DWORD __errCode) {
    std::fstream __log = std::fstream("log.txt", std::fstream::app);
    if (!__log.is_open()) { return; }
	__log << "ERROR: " << __msg
		<< "; error code: " << std::hex << __errCode;
	__log.close();
}