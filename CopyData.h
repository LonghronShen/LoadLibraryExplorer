// You can #define this value before doing an include to get a larger value
#ifndef MAX_COPY_LENGTH
#define MAX_COPY_LENGTH (MAX_PATH*sizeof(TCHAR))
#endif

class CopyPacket {
    public:
       GUID Signature;
       BYTE data[MAX_COPY_LENGTH];
    public:
       BOOL Same(const GUID & s)
             { return memcmp(&Signature, &s, sizeof(GUID)) == 0; }
};


class CopyData {
    public:
       CopyData(UINT id, GUID s)
             { packet.Signature = s; cds.dwData = id; cds.cbData = 0; cds.lpData = &packet; }
       UINT SetLength(UINT n)
             { cds.cbData = sizeof(packet.Signature) + n; return cds.cbData; }
       void SetData(LPCVOID src, size_t length)
             { ::CopyMemory(packet.data, src, length); }
       LRESULT Send(HWND target, HWND sender)
             { return ::SendMessage(target, WM_COPYDATA, (WPARAM)sender, (LPARAM)&cds); }
       LPBYTE GetData() { return packet.data; }
       static UINT GetMinimumLength()
             { return sizeof(GUID); }
    protected:
       COPYDATASTRUCT cds;
       CopyPacket packet;
};
