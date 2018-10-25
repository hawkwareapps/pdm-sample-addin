using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPDM.Interop.epdm;

namespace PDMAddin
{
    public class Addin : IEdmAddIn5
    {
        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            poInfo.mbsAddInName = "Custom Addin";
            poInfo.mbsCompany = "Hawk Ridge Systems";
            poInfo.mbsDescription = "My custom addin";
            poInfo.mlAddInVersion = 1;
            poInfo.mlRequiredVersionMajor = 18;
            poInfo.mlRequiredVersionMinor = 0;

            // create commands
            // create context menu commands
            poCmdMgr.AddCmd(100, "CustomAddin\\Action1");
            poCmdMgr.AddCmd(200, "CustomAddin\\Assy Only BOM");
        }

        public void OnCmd(ref EdmCmd poCmd, ref EdmCmdData[] ppoData)
        {
            IEdmVault10 edmVault = (IEdmVault10)poCmd.mpoVault;

            switch (poCmd.meCmdType) {
                case EdmCmdType.EdmCmd_Menu:
                    // match the CmdID with the command ID specified above
                    switch (poCmd.mlCmdID) {
                        case 100:
                            // CustomAddin -> Action1
                            string message = "";

                            for (int i = 0; i < ppoData.Length; i++) {
                                message += ppoData[i].mbsStrData1 + "\n";
                            }

                            edmVault.MsgBox(0, message);
                            break;
                        case 200:
                            // CustomAddin -> Assy Only BOM
                            for (int i = 0; i < ppoData.Length; i++) {
                                IEdmFile10 file = (IEdmFile10)edmVault.GetObject(EdmObjectType.EdmObject_File, ppoData[i].mlObjectID1);

                                IEdmBomView3 bomView = (IEdmBomView3)file.GetComputedBOM("BOM", 0, "@", (int)EdmBomFlag.EdmBf_AsBuilt | (int)EdmBomFlag.EdmBf_ShowSelected);

                                string tempFileName = System.IO.Path.GetTempFileName();
                                bomView.SaveToCSV(tempFileName, true);

                                string[] lines = System.IO.File.ReadAllLines(tempFileName);

                                string csvFile = System.IO.Path.GetTempFileName() + ".csv";

                                using(System.IO.TextWriter writer = new System.IO.StreamWriter(csvFile)) {
                                    for (int j = 0; j < lines.Length; j++) {
                                        if(j == 0) {
                                            writer.WriteLine(lines[j]);
                                        } else if (lines[j].ToUpper().Contains("SLDASM")) {
                                            writer.WriteLine(lines[j]);
                                        }
                                    }

                                    System.Diagnostics.Process.Start(csvFile);
                                }
                            }
                            break;
                        default:
                            break;
                    }
                    break;
            }
        }
    }
}
