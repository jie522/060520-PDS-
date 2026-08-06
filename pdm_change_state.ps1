# PDM 變更狀態 helper（由 app.py 以 subprocess 呼叫）
# 為什麼需要這支：Python 的 IDispatch 呼叫 IEdmFile5.ChangeState 一律回
# DISP_E_MEMBERNOTFOUND（找不到成員），必須走 .NET Interop 的 vtable 早期繫結；
# 且「00-提出申請」轉換設了身分驗證，要用 ChangeState3 帶密碼。
# 參數一律走環境變數（避免密碼出現在命令列/工作管理員）：
#   PDM_CS_FILE / PDM_CS_STATE / PDM_CS_TRANSITION / PDM_CS_COMMENT / PDM_CS_PASSWORD
# 輸出（stdout 最後一行）：OK: <狀態> | STILL: <狀態> | ERROR: <訊息>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

try {
    [Reflection.Assembly]::LoadFrom('C:\Program Files\SOLIDWORKS PDM\EPDM.Interop.epdm.dll') | Out-Null
    $src = @'
using EPDM.Interop.epdm;
using System.Threading;

public static class PdmCs
{
    public static string Run(string filePath, string target, string transition, string comment, string password)
    {
        IEdmVault5 vault = new EdmVault5Class();
        vault.LoginAuto("MAXCLAW", 0);
        IEdmFolder5 folder;
        IEdmFile5 file5 = vault.GetFileFromPath(filePath, out folder);
        if (file5 == null) return "ERROR: 找不到檔案";
        IEdmFile17 file = (IEdmFile17)file5;
        if (file.CurrentState.Name == target) return "OK: " + target;
        object state = target;
        object trans = transition;
        file.ChangeState3(ref state, ref trans, folder.ID, comment, 0, 0, password);
        for (int i = 0; i < 6; i++)
        {
            Thread.Sleep(2000);
            file.Refresh();
            if (file.CurrentState.Name == target) return "OK: " + file.CurrentState.Name;
        }
        return "STILL: " + file.CurrentState.Name;
    }
}
'@
    Add-Type -TypeDefinition $src -ReferencedAssemblies 'C:\Program Files\SOLIDWORKS PDM\EPDM.Interop.epdm.dll'
    $r = [PdmCs]::Run($env:PDM_CS_FILE, $env:PDM_CS_STATE, $env:PDM_CS_TRANSITION, $env:PDM_CS_COMMENT, $env:PDM_CS_PASSWORD)
    Write-Output $r
} catch {
    Write-Output ("ERROR: " + $_.Exception.Message)
}
