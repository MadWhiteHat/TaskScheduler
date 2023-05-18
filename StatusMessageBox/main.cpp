#include <iostream>
#include <string>

#include <windows.h>

int wmain(const int argc, const wchar_t** argv) {

	std::wstring __args;
	for (int i = 1; i < argc; ++i) {
		__args += argv[i];
		__args += L" ";
	}
	if (!__args.empty()) { __args.pop_back(); }

	MessageBoxW(
		NULL,
		__args.data(),
		L"Event triggered",
		MB_OK | MB_ICONINFORMATION
	);
}