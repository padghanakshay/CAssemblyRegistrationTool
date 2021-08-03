# CAssemblyRegistrationTool
This software use to register C# Class same as regasm.exe

If you want to use this using cpp use follwing code.

//-----------------------------------------------------------------------------
void OpenExe(LPCTSTR lpApplicationName, const std::wstring& command)
{
    // additional information
    STARTUPINFO si;     
    PROCESS_INFORMATION pi;

    // set the size of the structures
    ZeroMemory( &si, sizeof(si) );
    si.cb = sizeof(si);
    ZeroMemory( &pi, sizeof(pi) );
    LPWSTR pCommand = const_cast<LPWSTR>(command.c_str());
    // start the program up
    CreateProcess( lpApplicationName,   // the path
        pCommand,        // Command line
        NULL,           // Process handle not inheritable
        NULL,           // Thread handle not inheritable
        FALSE,          // Set handle inheritance to FALSE
        0,              // No creation flags
        NULL,           // Use parent's environment block
        NULL,           // Use parent's starting directory 
        &si,            // Pointer to STARTUPINFO structure
        &pi             // Pointer to PROCESS_INFORMATION structure (removed extra parentheses)
        );
    // Close process and thread handles. 
    CloseHandle( pi.hProcess );
    CloseHandle( pi.hThread );
}

OpenExe(L"Full path of CAssemblyRegistrationTool.exe", L"Full path of assembly file like As.dll");
