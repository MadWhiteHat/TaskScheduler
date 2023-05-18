#include <iostream>
#include <fstream>
#include <vector>

#include <windows.h>
#include <comdef.h>
#include <taskschd.h>
#include <wincred.h>

#include "create_task_scheduler.h"

BOOL
ITask::CreateTask(
    ITaskService** __service,
    ITaskDefinition** __task,
    LPCWSTR __authorName
) {
    HRESULT __hr = S_OK;

    IRegistrationInfo* __taskRegInfo = nullptr;
    ITaskSettings* __taskSettings = nullptr;

    if (__authorName == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("Task's author name not set", 0);
        return FALSE;
    }

    __hr = (*__service)->NewTask(0, __task);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("ITaskService::NewTask", __hr);
        return FALSE;
    }

    __hr = (*__task)->get_RegistrationInfo(&__taskRegInfo);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("ITaskDefinition::get_RegistrationInfo", __hr);
        return FALSE;
    }

    __hr = __taskRegInfo->put_Author(_bstr_t(__authorName));
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&*__taskRegInfo));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("IRegistrationInfo::put_Author", __hr);
        return FALSE;
    }

    __hr = (*__task)->get_Settings(&__taskSettings);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&*__taskRegInfo));
        IRelease(reinterpret_cast<IUnknown**>(&__taskSettings));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("ITaskDefinition::get_Settings", __hr);
        return FALSE;
    }

    __hr = __taskSettings->put_StartWhenAvailable(VARIANT_TRUE);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&*__taskRegInfo));
        IRelease(reinterpret_cast<IUnknown**>(&__taskSettings));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("ITaskSettings::put_StartWhenAvailable", __hr);
        return FALSE;
    }

    IRelease(reinterpret_cast<IUnknown**>(&__taskRegInfo));
    IRelease(reinterpret_cast<IUnknown**>(&__taskSettings));
    return TRUE;
}

BOOL
ITask::AddTriggerEvent(
    ITaskService** __service,
    ITaskDefinition** __task,
    LPCWSTR __triggerDelay,
    LPCWSTR __triggerName,
    LPCWSTR __triggerQueryString,
    std::vector<ITask::EventData>& __triggerValueQueries
) {
    HRESULT __hr = S_OK;

    ITriggerCollection* __taskTriggerCollection = nullptr;
    ITrigger* __taskTrigger = nullptr;
    IEventTrigger* __taskEventTrigger = nullptr;
    ITaskNamedValueCollection* __taskNamedValCollection = nullptr;
    ITaskNamedValuePair* __taskNamedValPair = nullptr;

    if ((*__service) == nullptr || (*__task) == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("Task Scheduler init failure", 0);
        return FALSE;
    }

    if (__triggerName == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("Trigger name not set", 0);
        return FALSE;
    }

    if (__triggerQueryString == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("Trigger query string not set", 0);
        return FALSE;
    }

    if (__triggerValueQueries.empty()) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("Trigger's array of name-value pairs is empty", 0);
        return FALSE;
    }

    __hr = (*__task)->get_Triggers(&__taskTriggerCollection);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("ITaskDefinition::get_Triggers", __hr);
        return FALSE;
    }

    __hr = __taskTriggerCollection->Create(TASK_TRIGGER_EVENT, &__taskTrigger);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("ITriggerCollection::Create", __hr);
        return FALSE;
    }

    __hr = __taskTrigger->QueryInterface(
        IID_IEventTrigger,
        reinterpret_cast<void**>(&__taskEventTrigger));
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskTrigger));
        IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("ITrigger::QueryEventInterface", __hr);
        return FALSE;
    }

    if (__triggerDelay != nullptr) {
        __hr = __taskEventTrigger->put_Delay(_bstr_t(__triggerDelay));
        if (FAILED(__hr)) {
            IRelease(reinterpret_cast<IUnknown**>(&__taskEventTrigger));
            IRelease(reinterpret_cast<IUnknown**>(&__taskTrigger));
            IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));
            IRelease(reinterpret_cast<IUnknown**>(__task));
            IRelease(reinterpret_cast<IUnknown**>(__service));
            _PrintError("IEventTrigger::put_Delay", __hr);
            return FALSE;
        }
    }
    __hr = __taskEventTrigger->put_Id(_bstr_t(__triggerName));
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskEventTrigger));
        IRelease(reinterpret_cast<IUnknown**>(&__taskTrigger));
        IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("IEventTrigger::put_Id", __hr);
        return FALSE;
    }

    __hr = __taskEventTrigger->put_Subscription(_bstr_t(__triggerQueryString));
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskEventTrigger));
        IRelease(reinterpret_cast<IUnknown**>(&__taskTrigger));
        IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("IEventTrigger::put_Subscription", __hr);
        return FALSE;
    }

    __hr = __taskEventTrigger->get_ValueQueries(&__taskNamedValCollection);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskEventTrigger));
        IRelease(reinterpret_cast<IUnknown**>(&__taskTrigger));
        IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        IRelease(reinterpret_cast<IUnknown**>(__service));
        _PrintError("IEventTrigger::get_ValueQueries", __hr);
        return FALSE;
    }

    for (const auto& __el : __triggerValueQueries) {
        __hr = __taskNamedValCollection->Create(
            _bstr_t(__el._name),
            _bstr_t(__el._queryString),
            &__taskNamedValPair
        );

        if (FAILED(__hr)) {
            IRelease(reinterpret_cast<IUnknown**>(&__taskNamedValCollection));
            IRelease(reinterpret_cast<IUnknown**>(&__taskEventTrigger));
            IRelease(reinterpret_cast<IUnknown**>(&__taskTrigger));
            IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));
            IRelease(reinterpret_cast<IUnknown**>(__task));
            IRelease(reinterpret_cast<IUnknown**>(__service));
            _PrintError("ITaskNamedValueCollection::Create", __hr);
            return FALSE;
        }

        IRelease(reinterpret_cast<IUnknown**>(&__taskNamedValPair));
    }

    IRelease(reinterpret_cast<IUnknown**>(&__taskNamedValCollection));
    IRelease(reinterpret_cast<IUnknown**>(&__taskEventTrigger));
    IRelease(reinterpret_cast<IUnknown**>(&__taskTrigger));
    IRelease(reinterpret_cast<IUnknown**>(&__taskTriggerCollection));

    return TRUE;
}

BOOL
ITask::AddExecAction(
    ITaskDefinition** __task,
    LPCWSTR __taskExecName,
    std::vector<ITask::EventData>& __actionValueQueries
) {
    HRESULT __hr = S_OK;

    IActionCollection* __taskActionCollection = nullptr;
    IAction* __taskAction = nullptr;
    IExecAction* __taskExecAction = nullptr;

    if ((*__task) == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("Task Scheduler init failure", 0);
        return FALSE;
    }

    if (__taskExecName == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("Execution path not set", 0);
        return FALSE;
    }

    __hr = (*__task)->get_Actions(&__taskActionCollection);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("ITaskDefinition::get_Actions", __hr);
        return FALSE;
    }

    __hr = __taskActionCollection->Create(TASK_ACTION_EXEC, &__taskAction);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskActionCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("IActionCollection::Create", __hr);
        return FALSE;
    }

    __hr = __taskAction->QueryInterface(
        IID_IExecAction,
        reinterpret_cast<void**>(&__taskExecAction)
    );
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskAction));
        IRelease(reinterpret_cast<IUnknown**>(&__taskActionCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("IAction::QueryInterface", __hr);
        return FALSE;
    }

    __hr = __taskExecAction->put_Path(_bstr_t(__taskExecName));
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskExecAction));
        IRelease(reinterpret_cast<IUnknown**>(&__taskAction));
        IRelease(reinterpret_cast<IUnknown**>(&__taskActionCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("IExecAction::put_Path", __hr);
        return FALSE;
    }

    std::wstring __tmpArgs;

    for (const auto& __el : __actionValueQueries) {
        __tmpArgs += L"$(";
        __tmpArgs += __el._name;
        __tmpArgs += L") ";
    }
    if (!__tmpArgs.empty()) { __tmpArgs.pop_back(); }

    __hr = __taskExecAction->put_Arguments(_bstr_t(__tmpArgs.data()));
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskExecAction));
        IRelease(reinterpret_cast<IUnknown**>(&__taskAction));
        IRelease(reinterpret_cast<IUnknown**>(&__taskActionCollection));
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("IExecAction::put_Path", __hr);
        return FALSE;
    }

    IRelease(reinterpret_cast<IUnknown**>(&__taskExecAction));
    IRelease(reinterpret_cast<IUnknown**>(&__taskAction));
    IRelease(reinterpret_cast<IUnknown**>(&__taskActionCollection));
    return TRUE;
}

BOOL
ITask::RegisterTask(
    ITaskService** __service,
    ITaskDefinition** __task,
    LPCWSTR __folder,
    LPCWSTR __taskName
) {
    HRESULT __hr = S_OK;

    LPCWSTR __realFolder = (__folder == nullptr) ? L"\\" : __folder;
    IPrincipal* __taskPrincipal = nullptr;
    ITaskFolder* __taskFolder = nullptr;
    IRegisteredTask* __taskReg = nullptr;

    if ((*__service) == nullptr || (*__task) == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("Task Scheduler init failure", 0);
        return FALSE;
    }

    if (__taskName == nullptr) {
        IRelease(reinterpret_cast<IUnknown**>(__task));
        _PrintError("Task name not set", 0);
        return FALSE;
    }

    __hr = (*__task)->get_Principal(&__taskPrincipal);
    if (FAILED(__hr)) {
        _PrintError("ITaskDefinition::get_Principal", __hr);
        return FALSE;
    }

    __hr = __taskPrincipal->put_RunLevel(TASK_RUNLEVEL_HIGHEST);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskPrincipal));
        _PrintError("IPrincipal::put_RunLevel", __hr);
        return FALSE;
    }

    __hr = (*__service)->GetFolder(_bstr_t(__realFolder), &__taskFolder);
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskPrincipal));
        _PrintError("ITaskService::GetFolder", __hr);
        return FALSE;
    }

    __hr = __taskFolder->RegisterTaskDefinition(
        _bstr_t(__taskName),
        (*__task),
        TASK_CREATE_OR_UPDATE,
        _variant_t(),
        _variant_t(),
        TASK_LOGON_INTERACTIVE_TOKEN,
        _variant_t(L""),
        &__taskReg
    );
    if (FAILED(__hr)) {
        IRelease(reinterpret_cast<IUnknown**>(&__taskFolder));
        IRelease(reinterpret_cast<IUnknown**>(&__taskPrincipal));
        _PrintError("ITaskFolder::RegisterTaskDefinition", __hr);
        return FALSE;
    }

    std::wcout << "Task: " << __taskName << " was successfully registered!"
        << std::endl;

    IRelease(reinterpret_cast<IUnknown**>(&__taskReg));
    IRelease(reinterpret_cast<IUnknown**>(&__taskFolder));
    IRelease(reinterpret_cast<IUnknown**>(&__taskPrincipal));

    return TRUE;
}

BOOL
ITask::SchedulerInit(ITaskService** __service) {
    HRESULT __hr = S_OK;
    
    __hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(__hr)) {
        _PrintError("CoInitializeEx", __hr);
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
        _PrintError("CoInitializeSecurity", __hr);
        SchedulerDeinit(__service);
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
        _PrintError("CoCreateInstance", __hr);
        SchedulerDeinit(__service);
        return FALSE;
    }

    __hr = (*__service)->Connect(
        _variant_t(),
        _variant_t(),
        _variant_t(),
        _variant_t()
    );

    if (FAILED(__hr)) {
        _PrintError("ITaskService::Connect", __hr);
        IRelease(reinterpret_cast<IUnknown**>(__service));
        SchedulerDeinit(__service);
        return FALSE;
    }

    return TRUE;
}

void
ITask::SchedulerDeinit(ITaskService** __service) {
    IRelease(reinterpret_cast<IUnknown**>(__service));
    CoUninitialize();
}

void
ITask::IRelease(IUnknown** __iface) {
    if ((*__iface) != nullptr) {
        (*__iface)->Release();
        (*__iface) = nullptr;
    }
}

void
ITask::_PrintError(const char* __msg, DWORD __errCode) {
    std::fstream __log = std::fstream("log.txt", std::fstream::app);
    if (!__log.is_open()) { return; }
    __log << "ERROR: 0x" << std::hex << __errCode << "; " << __msg << std::endl;
    __log.close();
}
