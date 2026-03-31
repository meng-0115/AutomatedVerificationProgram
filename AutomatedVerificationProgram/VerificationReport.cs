using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Spire.Pdf.Exporting.XPS.Schema.Mc;

namespace AutomatedVerificationProgram
{
    internal class VerificationReport
    {
        public enum withRuleItems
        {
            Pb = 0,
            Cd = 1,
            Hg = 2,
            CrVI = 3,
            PPBs = 13,
            PPBEs = 23,
            DIBP = 24,
            DEHP = 25,
            BBP = 26,
            DBP = 27,
            Cl = 28,
            Br = 29,
            F = 30,
            PFOS = 31,
            PFOA = 32,
            Sb = 33,
            Be = 34,
            PVC = 35
        }
        internal class VerificationResultMatrix
        {
            String[] specItems = new String[]
            {
               "Lead", "Cadmium", "Mercury", "Hexavalent Chromium",
               "Monobromobiphenyl","Dibromobiphenyl","Tribromobiphenyl","Tetrabromobiphenyl","Pentabromobiphenyl","Hexabromobiphenyl","Heptabromobiphenyl","Octabromobiphenyl","Nonabromobiphenyl","Decabromobiphenyl",
               "Monobromodiphenyl ether","Dibromodiphenyl ether","Tribromodiphenyl ether","Tetrabromodiphenyl ether","Pentabromodiphenyl ether","Hexabromodiphenyl ether","Heptabromodiphenyl ether","Octabromodiphenyl ether","Nonabromodiphenyl ether","Decabromodiphenyl ether",
               "DIBP","DEHP","BBP","DBP",
               "Chlorine","Bromine","Fluorine","PFOS",
               "PFOA","Antimony","Beryllium","PVC",
               "PFAS"
            };
        }

        public class SpecItem
        {
            public string Name { get; set; }
            public float Metal { get; set; }
            public float NonMetal { get; set; }
            public bool MustND { get; set; }
            public string testMethod1 { get; set; }
            public string testMethod2 { get; set; }
            public float MdlValue { get; set; }
        }
        private List<SpecItem> specsStandards;

        #region (Cd)(Hg)(Xbromobiphenyl)(Xbromodiphenyl ether)(DIBP)(DEHP)(BBP)(DBP)(F)(PFOS)(PFOA)(Sb)
        private void InitializeSpecs()
        {
            // 1. 定義規格結構與資料表 
            
            specsStandards = new List<SpecItem>
            {
                new SpecItem{ Name = "Lead",                     Metal = 500f, NonMetal = 50f,  MustND = false, testMethod1 = "IEC 62321-5",    testMethod2 = "",             MdlValue = 2f  },
                new SpecItem{ Name = "Cadmium",                  Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-5",    testMethod2 = "",             MdlValue = 2f  },
                new SpecItem{ Name = "Mercury",                  Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-4",    testMethod2 = "",             MdlValue = 2f  },
                new SpecItem{ Name = "Hexavalent Chromium",      Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-7-1",  testMethod2 = "IEC 62321-7-2",             MdlValue = 0.1f },

                //Xbromobiphenyl  4~13
                new SpecItem{ Name = "Monobromobiphenyl",        Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Dibromobiphenyl",          Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Tribromobiphenyl",         Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Tetrabromobiphenyl",       Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Pentabromobiphenyl",       Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Hexabromobiphenyl",        Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Heptabromobiphenyl",       Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Octabromobiphenyl",        Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Nonabromobiphenyl",        Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Decabromobiphenyl",        Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                    
                //Xbromodiphenyl ether 14~23
                new SpecItem{ Name = "Monobromodiphenyl ether",  Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Dibromodiphenyl ether",    Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Tribromodiphenyl ether",   Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Tetrabromodiphenyl ether", Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Pentabromodiphenyl ether", Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Hexabromodiphenyl ether",  Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Heptabromodiphenyl ether", Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Octabromodiphenyl ether",  Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Nonabromodiphenyl ether",  Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },
                new SpecItem{ Name = "Decabromodiphenyl ether",  Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-6", testMethod2 = "",             MdlValue = 5f  },

                new SpecItem{ Name = "DIBP",                     Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-8", testMethod2 = "",             MdlValue = 50f },
                new SpecItem{ Name = "DEHP",                     Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-8", testMethod2 = "",             MdlValue = 50f },
                new SpecItem{ Name = "BBP",                      Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-8", testMethod2 = "",             MdlValue = 50f },
                new SpecItem{ Name = "DBP",                      Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "IEC 62321-8", testMethod2 = "",             MdlValue = 50f },

                new SpecItem{ Name = "Chlorine",                 Metal = 0f,   NonMetal = 700f,   MustND = true,  testMethod1 = "EN 14582",    testMethod2 = "",             MdlValue = 50f },
                new SpecItem{ Name = "Bromine",                  Metal = 0f,   NonMetal = 700f,   MustND = true,  testMethod1 = "EN 14582",    testMethod2 = "",             MdlValue = 50f },

                new SpecItem{ Name = "Fluorine",                 Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "EN 14582",    testMethod2 = "",             MdlValue = 50f },
                new SpecItem{ Name = "PFOS",                     Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "EN 17681-1",  testMethod2 = "EN 17681-2",   MdlValue = 25f },
                new SpecItem{ Name = "PFOA",                     Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "EN 17681-1",  testMethod2 = "EN 17681-2",   MdlValue = 25f },

                new SpecItem{ Name = "Antimony",                 Metal = 300f, NonMetal = 300f,   MustND = false, testMethod1 = "EPA 3052",    testMethod2 = "EPA 3050B",    MdlValue = 10f },
                new SpecItem{ Name = "Beryllium",                Metal = 100f, NonMetal = 0f,   MustND = true,  testMethod1 = "EPA 3052",    testMethod2 = "EPA 3050B",    MdlValue = 10f },
                new SpecItem{ Name = "PFAS",                     Metal = 0f,   NonMetal = 0f,   MustND = true,  testMethod1 = "",            testMethod2 = "",             MdlValue = 10f }
                };
        }
        #endregion


        //驗證報告有三個參數：審核情境、金屬非金屬、均勻材質
        public List<string> verificationReport(int VerificationScenarioIndex, int MetalNonmetalIndex, ReportRowDataManager manager, ReportParaments reportParaments = null, string homogeneousMaterialType = null, int tpiCompany = -1, bool autoVertificationFlag = false, string supplierNumer = "", string homogeneousNumber = "", int pbContain = 0, int halogenContain = 0)
        {
            List<string> resultList = new List<string>();
            InitializeSpecs();

            int verificationItemCount = 0;

            bool PBBsTested = false;
            bool PBBsMethod = false;
            bool PBBsResult = false;
            bool PBBsMdl = false;

            bool PBBEsTested = false;
            bool PBBEsMethod = false;
            bool PBBEsResult = false;
            bool PBBEsMdl = false;

            float clPlusBrResultVal = 0;

            #region 1-a
            if (VerificationScenarioIndex == 0)
            {
                if(reportParaments != null)
                {
                    resultList.Add($"報告號碼: {reportParaments.reportNunber}\r\n");
                    resultList.Add($"廠商名稱: {reportParaments.company}\r\n");
                    resultList.Add($"測試期間的最後日期: {reportParaments.finalExtractedDate}\r\n");
                    string status = DateTime.Parse(reportParaments.finalExtractedDate).AddYears(1) > DateTime.Now ? "有效" : "過期";
                    resultList.Add($"效期: {status}\r\n");
                    resultList.Add($"Sample Name: {reportParaments.sampleName}\r\n");
                    resultList.Add($"Style Item No.: {reportParaments.styleItemNo}\r\n");
                    resultList.Add($"樣品描述: {reportParaments.testParts}\r\n\r\n");
                }

                //所有檢查項目列表
                string[] specItems = new[] {
                    "Lead", "Cadmium", "Mercury", "Hexavalent Chromium",
                    "Monobromobiphenyl","Dibromobiphenyl","Tribromobiphenyl","Tetrabromobiphenyl","Pentabromobiphenyl","Hexabromobiphenyl","Heptabromobiphenyl","Octabromobiphenyl","Nonabromobiphenyl","Decabromobiphenyl",
                    "Monobromodiphenyl ether","Dibromodiphenyl ether","Tribromodiphenyl ether","Tetrabromodiphenyl ether","Pentabromodiphenyl ether","Hexabromodiphenyl ether","Heptabromodiphenyl ether","Octabromodiphenyl ether","Nonabromodiphenyl ether","Decabromodiphenyl ether",
                    "DIBP","DEHP","BBP","DBP",
                    "Chlorine","Bromine","Fluorine","PFOS",
                    "PFOA","Antimony","Beryllium","PVC",
                    "PFAS"
                };

                
                

                foreach (var s in specsStandards)
                {
                    bool vertificationMethod = false;
                    bool vertificationResult = false;
                    bool vertificationMdl = false;

                    float vertificationResultVal = 0;
                    float vertificationMdlVal = 0;

                   
                    //沒有檢測的直接跳過並記錄在報告裡面
                    var row = manager.reportRowData.FirstOrDefault(r => r.TestItem.Contains(specItems[verificationItemCount]));
                    
                    /*---------------------------------------------------------------------*/
                    #region 檢測結果數值轉換
                    if (row.Result.ToLower() == "n.d.")
                    {
                        vertificationResultVal = 0;
                    }
                    else if (float.TryParse(row.Result, out float parsedVal))
                    {
                        vertificationResultVal = parsedVal;
                    }
                    #endregion
                    /*---------------------------------------------------------------------*/
                    #region (Pb) 編號0
                    if (verificationItemCount == 0)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Pb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Pb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "wire" || homogeneousMaterialType == "lead frame")
                        {
                            if (vertificationResultVal < s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal < s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(pb)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (tpiCompany == 5)
                            {
                                if (vertificationMdlVal < 5)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if(!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        { 
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Pb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Pb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Pb"));
                            }
                        }
                        #endregion


                    }
                    #endregion
                    #region Cr(VI) 編號3
                    else if (verificationItemCount == 3)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Cr(VI): 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Cr(VI)"));
                            }
                            verificationItemCount++;
                            continue;
                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (homogeneousMaterialType == "Immersion Tin" || homogeneousMaterialType == "Plating/Coatings layer" ||
                           homogeneousMaterialType == "Solder paste")
                        {
                            if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (row.Method.Contains(s.testMethod1))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1)
                        {
                            if (row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (MetalNonmetalIndex == 0)
                            {
                                if (vertificationMdlVal == 0.1 || vertificationMdlVal < 0.1)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (MetalNonmetalIndex == 1)
                            {
                                if (vertificationMdlVal <= 8)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Cr(VI)"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Cr(VI)"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Cr(VI)"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PPBs
                    else if (verificationItemCount >= 4 && verificationItemCount <= 13)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBsTested = false;
                        }
                        else
                        {
                            PBBsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBsResult = true;
                        }
                        else
                        {
                            PBBsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBsMethod = true;
                        }
                        else
                        {
                            PBBsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBsMdl = true;
                            }
                            else
                            {
                                PBBsMdl = false;
                            }
                        }
                        else
                        {
                            PBBsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if(PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBs"));
                                }

                                if (PBBsResult && PBBsMethod && !PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBs"));
                                }

                                if (!PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBs"));
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion
                    #region PPBEs
                    else if (verificationItemCount >= 14 && verificationItemCount <= 23)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBEsTested = false;
                        }
                        else
                        {
                            PBBEsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBEsResult = true;
                        }
                        else
                        {
                            PBBEsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBEsMethod = true;
                        }
                        else
                        {
                            PBBEsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBEsMdl = true;
                            }
                            else
                            {
                                PBBEsMdl = false;
                            }
                        }
                        else
                        {
                            PBBEsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBEs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBEs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBEs"));
                                }

                                if (PBBEsResult && PBBEsMethod && !PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBEs"));
                                }

                                if (!PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBEs"));
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Cl)(Br)
                    else if (verificationItemCount == 28 || verificationItemCount == 29)
                    {
                        string elementName = verificationItemCount == 28 ? "Cl" : "Br";

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Solder paste")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal == s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }

                        clPlusBrResultVal += vertificationResultVal;
                        

                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成                       
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (F)
                    else if (verificationItemCount == 30)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"F: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "F"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        vertificationResult = true;
                        /*F的檢測結果不論數值為何，只要有檢測到就透過outllook發郵件通知相關人員*/
                        if (vertificationResultVal != 0 && autoVertificationFlag == true)
                        {
                            Mail.Outlook_to_F(supplierNumer, homogeneousNumber, row.Result);
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "F"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "F"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "F"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PFOS/PFOA
                    else if (verificationItemCount == 31 || verificationItemCount == 32)
                    {
                        string elementName = verificationItemCount == 31 ? "PFOS" : "PFOA";
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Sb)
                    else if (verificationItemCount == 33)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Sb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Sb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal <= s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if(!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        { 
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Sb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Sb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Sb"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Be)
                    else if (verificationItemCount == 34)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Be: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Be"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定   
                        if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal <= s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Beryllium copper")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(Be)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }

                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Be"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Be"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Be"));
                            }
                        }
                        #endregion
                    }
                    #endregion

                    #region (PVC)                    
                    else if (verificationItemCount == 35)
                    {
                        /*
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PVC: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PVC"));
                            }
                            verificationItemCount++;
                            continue;

                        }

                        #region 檢測結果判定    
                        vertificationResult = true;
                        #endregion
                       
                        #region 檢測方法判定
                        vertificationMethod = true;
                        #endregion
                       
                        #region mdl標準檢測
                        vertificationMethod = true;
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PVC"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PVC"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PVC"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion

                    #region 其他有規則
                    else
                    {
                        if(verificationItemCount == 36)
                        {
                            break;
                        }

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", ((withRuleItems)verificationItemCount).ToString()));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", ((withRuleItems)verificationItemCount).ToString()));
                            }
                        }
                        #endregion
                    }
                    #endregion

                    verificationItemCount++;
                }

            }
            #endregion
            #region 1-b
            else if (VerificationScenarioIndex == 1)
            {
                if (reportParaments != null)
                {
                    resultList.Add($"報告號碼: {reportParaments.reportNunber}\r\n");
                    resultList.Add($"廠商名稱: {reportParaments.company}\r\n");
                    resultList.Add($"測試期間的最後日期: {reportParaments.finalExtractedDate}\r\n");
                    string status = DateTime.Parse(reportParaments.finalExtractedDate).AddYears(1) > DateTime.Now ? "有效" : "過期";
                    resultList.Add($"效期: {status}\r\n");
                    resultList.Add($"Sample Name: {reportParaments.sampleName}\r\n");
                    resultList.Add($"Style Item No.: {reportParaments.styleItemNo}\r\n");
                    resultList.Add($"樣品描述: {reportParaments.testParts}\r\n\r\n");
                }

                //所有檢查項目列表
                string[] specItems = new[] {
                    "Lead", "Cadmium", "Mercury", "Hexavalent Chromium",
                    "Monobromobiphenyl","Dibromobiphenyl","Tribromobiphenyl","Tetrabromobiphenyl","Pentabromobiphenyl","Hexabromobiphenyl","Heptabromobiphenyl","Octabromobiphenyl","Nonabromobiphenyl","Decabromobiphenyl",
                    "Monobromodiphenyl ether","Dibromodiphenyl ether","Tribromodiphenyl ether","Tetrabromodiphenyl ether","Pentabromodiphenyl ether","Hexabromodiphenyl ether","Heptabromodiphenyl ether","Octabromodiphenyl ether","Nonabromodiphenyl ether","Decabromodiphenyl ether",
                    "DIBP","DEHP","BBP","DBP",
                    "Chlorine","Bromine","Fluorine","PFOS",
                    "PFOA","Antimony","Beryllium","PVC",
                    "PFAS"
                };




                foreach (var s in specsStandards)
                {
                    bool vertificationMethod = false;
                    bool vertificationResult = false;
                    bool vertificationMdl = false;

                    float vertificationResultVal = 0;
                    float vertificationMdlVal = 0;


                    //沒有檢測的直接跳過並記錄在報告裡面
                    var row = manager.reportRowData.FirstOrDefault(r => r.TestItem.Contains(specItems[verificationItemCount]));

                    /*---------------------------------------------------------------------*/
                    #region 檢測結果數值轉換
                    if (row.Result.ToLower() == "n.d.")
                    {
                        vertificationResultVal = 0;
                    }
                    else if (float.TryParse(row.Result, out float parsedVal))
                    {
                        vertificationResultVal = parsedVal;
                    }
                    #endregion
                    /*---------------------------------------------------------------------*/
                    #region (Pb) 編號0
                    if (verificationItemCount == 0)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Pb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Pb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "wire" || homogeneousMaterialType == "lead frame")
                        {
                            if (vertificationResultVal < s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal < s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(pb)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (tpiCompany == 5)
                            {
                                if (vertificationMdlVal < 5)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Pb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Pb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Pb"));
                            }
                        }
                        #endregion


                    }
                    #endregion
                    #region Cr(VI) 編號3
                    else if (verificationItemCount == 3)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Cr(VI): 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Cr(VI)"));
                            }
                            verificationItemCount++;
                            continue;
                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (homogeneousMaterialType == "Immersion Tin" || homogeneousMaterialType == "Plating/Coatings layer" ||
                           homogeneousMaterialType == "Solder paste")
                        {
                            if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (row.Method.Contains(s.testMethod1))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1)
                        {
                            if (row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (MetalNonmetalIndex == 0)
                            {
                                if (vertificationMdlVal == 0.1 || vertificationMdlVal < 0.1)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (MetalNonmetalIndex == 1)
                            {
                                if (vertificationMdlVal <= 8)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Cr(VI)"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Cr(VI)"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Cr(VI)"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PPBs
                    else if (verificationItemCount >= 4 && verificationItemCount <= 13)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBsTested = false;
                        }
                        else
                        {
                            PBBsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBsResult = true;
                        }
                        else
                        {
                            PBBsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBsMethod = true;
                        }
                        else
                        {
                            PBBsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBsMdl = true;
                            }
                            else
                            {
                                PBBsMdl = false;
                            }
                        }
                        else
                        {
                            PBBsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBs"));
                                }

                                if (PBBsResult && PBBsMethod && !PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBs"));
                                }

                                if (!PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBs"));
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion
                    #region PPBEs
                    else if (verificationItemCount >= 14 && verificationItemCount <= 23)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBEsTested = false;
                        }
                        else
                        {
                            PBBEsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBEsResult = true;
                        }
                        else
                        {
                            PBBEsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBEsMethod = true;
                        }
                        else
                        {
                            PBBEsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBEsMdl = true;
                            }
                            else
                            {
                                PBBEsMdl = false;
                            }
                        }
                        else
                        {
                            PBBEsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBEs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBEs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBEs"));
                                }

                                if (PBBEsResult && PBBEsMethod && !PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBEs"));
                                }

                                if (!PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBEs"));
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Cl)(Br)
                    else if (verificationItemCount == 28 || verificationItemCount == 29)
                    {
                        string elementName = verificationItemCount == 28 ? "Cl" : "Br";

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Solder paste")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal == s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }

                        clPlusBrResultVal += vertificationResultVal;


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成                       
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (F)
                    else if (verificationItemCount == 30)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"F: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "F"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        vertificationResult = true;
                        /*F的檢測結果不論數值為何，只要有檢測到就透過outllook發郵件通知相關人員*/
                        if (vertificationResultVal != 0 && autoVertificationFlag == true)
                        {
                            Mail.Outlook_to_F(supplierNumer, homogeneousNumber, row.Result);
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "F"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "F"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "F"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PFOS/PFOA
                    else if (verificationItemCount == 31 || verificationItemCount == 32)
                    {
                        string elementName = verificationItemCount == 31 ? "PFOS" : "PFOA";
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Sb)
                    else if (verificationItemCount == 33)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Sb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Sb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal <= s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Sb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Sb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Sb"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Be)
                    else if (verificationItemCount == 34)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Be: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Be"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定   
                        if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal <= s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Beryllium copper")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(Be)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }

                        #endregion
                        
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Be"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Be"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Be"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (PVC)                    
                    else if (verificationItemCount == 35)
                    {
                        /*
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PVC: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PVC"));
                            }
                            verificationItemCount++;
                            continue;

                        }

                        #region 檢測結果判定    
                        vertificationResult = true;
                        #endregion
                       
                        #region 檢測方法判定
                        vertificationMethod = true;
                        #endregion
                       
                        #region mdl標準檢測
                        vertificationMethod = true;
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PVC"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PVC"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PVC"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region 其他有規則
                    else
                    {
                        if (verificationItemCount == 36)
                        {
                            break;
                        }

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", ((withRuleItems)verificationItemCount).ToString()));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", ((withRuleItems)verificationItemCount).ToString()));
                            }
                        }
                        #endregion
                    }
                    #endregion

                    verificationItemCount++;
                }
            }
            #endregion
            #region 2
            else if (VerificationScenarioIndex == 2)
            {
                if (reportParaments != null)
                {
                    resultList.Add($"報告號碼: {reportParaments.reportNunber}\r\n");
                    resultList.Add($"廠商名稱: {reportParaments.company}\r\n");
                    resultList.Add($"測試期間的最後日期: {reportParaments.finalExtractedDate}\r\n");
                    string status = DateTime.Parse(reportParaments.finalExtractedDate).AddYears(1) > DateTime.Now ? "有效" : "過期";
                    resultList.Add($"效期: {status}\r\n");
                    resultList.Add($"Sample Name: {reportParaments.sampleName}\r\n");
                    resultList.Add($"Style Item No.: {reportParaments.styleItemNo}\r\n");
                    resultList.Add($"樣品描述: {reportParaments.testParts}\r\n\r\n");
                }

                //所有檢查項目列表
                string[] specItems = new[] {
                    "Lead", "Cadmium", "Mercury", "Hexavalent Chromium",
                    "Monobromobiphenyl","Dibromobiphenyl","Tribromobiphenyl","Tetrabromobiphenyl","Pentabromobiphenyl","Hexabromobiphenyl","Heptabromobiphenyl","Octabromobiphenyl","Nonabromobiphenyl","Decabromobiphenyl",
                    "Monobromodiphenyl ether","Dibromodiphenyl ether","Tribromodiphenyl ether","Tetrabromodiphenyl ether","Pentabromodiphenyl ether","Hexabromodiphenyl ether","Heptabromodiphenyl ether","Octabromodiphenyl ether","Nonabromodiphenyl ether","Decabromodiphenyl ether",
                    "DIBP","DEHP","BBP","DBP",
                    "Chlorine","Bromine","Fluorine","PFOS",
                    "PFOA","Antimony","Beryllium","PVC",
                    "PFAS"
                };




                foreach (var s in specsStandards)
                {
                    bool vertificationMethod = false;
                    bool vertificationResult = false;
                    bool vertificationMdl = false;

                    float vertificationResultVal = 0;
                    float vertificationMdlVal = 0;


                    //沒有檢測的直接跳過並記錄在報告裡面
                    var row = manager.reportRowData.FirstOrDefault(r => r.TestItem.Contains(specItems[verificationItemCount]));

                    /*---------------------------------------------------------------------*/
                    #region 檢測結果數值轉換
                    if (row.Result.ToLower() == "n.d.")
                    {
                        vertificationResultVal = 0;
                    }
                    else if (float.TryParse(row.Result, out float parsedVal))
                    {
                        vertificationResultVal = parsedVal;
                    }
                    #endregion
                    /*---------------------------------------------------------------------*/
                    #region (Pb) 編號0
                    if (verificationItemCount == 0)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Pb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Pb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "wire" || homogeneousMaterialType == "lead frame")
                        {
                            if (vertificationResultVal < s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal < s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(pb)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (tpiCompany == 5)
                            {
                                if (vertificationMdlVal < 5)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Pb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Pb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Pb"));
                            }
                        }
                        #endregion


                    }
                    #endregion
                    #region Cr(VI) 編號3
                    else if (verificationItemCount == 3)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Cr(VI): 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Cr(VI)"));
                            }
                            verificationItemCount++;
                            continue;
                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (homogeneousMaterialType == "Immersion Tin" || homogeneousMaterialType == "Plating/Coatings layer" ||
                           homogeneousMaterialType == "Solder paste")
                        {
                            if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (row.Method.Contains(s.testMethod1))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1)
                        {
                            if (row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (MetalNonmetalIndex == 0)
                            {
                                if (vertificationMdlVal == 0.1 || vertificationMdlVal < 0.1)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (MetalNonmetalIndex == 1)
                            {
                                if (vertificationMdlVal <= 8)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Cr(VI)"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Cr(VI)"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Cr(VI)"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PPBs
                    else if (verificationItemCount >= 4 && verificationItemCount <= 13)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBsTested = false;
                        }
                        else
                        {
                            PBBsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBsResult = true;
                        }
                        else
                        {
                            PBBsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBsMethod = true;
                        }
                        else
                        {
                            PBBsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBsMdl = true;
                            }
                            else
                            {
                                PBBsMdl = false;
                            }
                        }
                        else
                        {
                            PBBsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBs"));
                                }

                                if (PBBsResult && PBBsMethod && !PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBs"));
                                }

                                if (!PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBs"));
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion
                    #region PPBEs
                    else if (verificationItemCount >= 14 && verificationItemCount <= 23)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBEsTested = false;
                        }
                        else
                        {
                            PBBEsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBEsResult = true;
                        }
                        else
                        {
                            PBBEsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBEsMethod = true;
                        }
                        else
                        {
                            PBBEsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBEsMdl = true;
                            }
                            else
                            {
                                PBBEsMdl = false;
                            }
                        }
                        else
                        {
                            PBBEsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBEs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBEs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBEs"));
                                }

                                if (PBBEsResult && PBBEsMethod && !PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBEs"));
                                }

                                if (!PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBEs"));
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Cl)(Br)
                    else if (verificationItemCount == 28 || verificationItemCount == 29)
                    {
                        string elementName = verificationItemCount == 28 ? "Cl" : "Br";

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Solder paste")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal == s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }

                        clPlusBrResultVal += vertificationResultVal;


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成                       
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (F)
                    else if (verificationItemCount == 30)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"F: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "F"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        vertificationResult = true;
                        /*F的檢測結果不論數值為何，只要有檢測到就透過outllook發郵件通知相關人員*/
                        if (vertificationResultVal != 0 && autoVertificationFlag == true)
                        {
                            Mail.Outlook_to_F(supplierNumer, homogeneousNumber, row.Result);
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "F"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "F"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "F"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PFOS/PFOA
                    else if (verificationItemCount == 31 || verificationItemCount == 32)
                    {
                        /*
                        string elementName = verificationItemCount == 31 ? "PFOS" : "PFOA";
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (Sb)
                    else if (verificationItemCount == 33)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Sb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Sb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal <= s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion

                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion

                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Sb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Sb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Sb"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (Be)
                    else if (verificationItemCount == 34)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Be: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Be"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定   
                        if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal <= s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Beryllium copper")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(Be)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }

                        #endregion
                        
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Be"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Be"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Be"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (PVC)                    
                    else if (verificationItemCount == 35)
                    {
                        /*
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PVC: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PVC"));
                            }
                            verificationItemCount++;
                            continue;

                        }

                        #region 檢測結果判定    
                        vertificationResult = true;
                        #endregion
                       
                        #region 檢測方法判定
                        vertificationMethod = true;
                        #endregion
                       
                        #region mdl標準檢測
                        vertificationMethod = true;
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PVC"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PVC"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PVC"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region 其他有規則
                    else
                    {
                        if (verificationItemCount == 36)
                        {
                            break;
                        }

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", ((withRuleItems)verificationItemCount).ToString()));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", ((withRuleItems)verificationItemCount).ToString()));
                            }
                        }
                        #endregion
                    }
                    #endregion

                    verificationItemCount++;
                }
            }
            #endregion
            #region 3a
            else if (VerificationScenarioIndex == 3)
            {
                if (reportParaments != null)
                {
                    resultList.Add($"報告號碼: {reportParaments.reportNunber}\r\n");
                    resultList.Add($"廠商名稱: {reportParaments.company}\r\n");
                    resultList.Add($"測試期間的最後日期: {reportParaments.finalExtractedDate}\r\n");
                    string status = DateTime.Parse(reportParaments.finalExtractedDate).AddYears(1) > DateTime.Now ? "有效" : "過期";
                    resultList.Add($"效期: {status}\r\n");
                    resultList.Add($"Sample Name: {reportParaments.sampleName}\r\n");
                    resultList.Add($"Style Item No.: {reportParaments.styleItemNo}\r\n");
                    resultList.Add($"樣品描述: {reportParaments.testParts}\r\n\r\n");
                }

                //所有檢查項目列表
                string[] specItems = new[] {
                    "Lead", "Cadmium", "Mercury", "Hexavalent Chromium",
                    "Monobromobiphenyl","Dibromobiphenyl","Tribromobiphenyl","Tetrabromobiphenyl","Pentabromobiphenyl","Hexabromobiphenyl","Heptabromobiphenyl","Octabromobiphenyl","Nonabromobiphenyl","Decabromobiphenyl",
                    "Monobromodiphenyl ether","Dibromodiphenyl ether","Tribromodiphenyl ether","Tetrabromodiphenyl ether","Pentabromodiphenyl ether","Hexabromodiphenyl ether","Heptabromodiphenyl ether","Octabromodiphenyl ether","Nonabromodiphenyl ether","Decabromodiphenyl ether",
                    "DIBP","DEHP","BBP","DBP",
                    "Chlorine","Bromine","Fluorine","PFOS",
                    "PFOA","Antimony","Beryllium","PVC",
                    "PFAS"
                };




                foreach (var s in specsStandards)
                {
                    bool vertificationMethod = false;
                    bool vertificationResult = false;
                    bool vertificationMdl = false;

                    float vertificationResultVal = 0;
                    float vertificationMdlVal = 0;


                    //沒有檢測的直接跳過並記錄在報告裡面
                    var row = manager.reportRowData.FirstOrDefault(r => r.TestItem.Contains(specItems[verificationItemCount]));

                    /*---------------------------------------------------------------------*/
                    #region 檢測結果數值轉換
                    if (row.Result.ToLower() == "n.d.")
                    {
                        vertificationResultVal = 0;
                    }
                    else if (float.TryParse(row.Result, out float parsedVal))
                    {
                        vertificationResultVal = parsedVal;
                    }
                    #endregion
                    /*---------------------------------------------------------------------*/
                    #region (Pb) 編號0
                    if (verificationItemCount == 0)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Pb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Pb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "wire" || homogeneousMaterialType == "lead frame")
                        {
                            if (vertificationResultVal < s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal < s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(pb)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (tpiCompany == 5)
                            {
                                if (vertificationMdlVal < 5)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Pb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Pb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Pb"));
                            }
                        }
                        #endregion


                    }
                    #endregion
                    #region Cr(VI) 編號3
                    else if (verificationItemCount == 3)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Cr(VI): 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Cr(VI)"));
                            }
                            verificationItemCount++;
                            continue;
                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (homogeneousMaterialType == "Immersion Tin" || homogeneousMaterialType == "Plating/Coatings layer" ||
                           homogeneousMaterialType == "Solder paste")
                        {
                            if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (row.Method.Contains(s.testMethod1))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1)
                        {
                            if (row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (MetalNonmetalIndex == 0)
                            {
                                if (vertificationMdlVal == 0.1 || vertificationMdlVal < 0.1)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (MetalNonmetalIndex == 1)
                            {
                                if (vertificationMdlVal <= 8)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Cr(VI)"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Cr(VI)"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Cr(VI)"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PPBs
                    else if (verificationItemCount >= 4 && verificationItemCount <= 13)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBsTested = false;
                        }
                        else
                        {
                            PBBsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBsResult = true;
                        }
                        else
                        {
                            PBBsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBsMethod = true;
                        }
                        else
                        {
                            PBBsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBsMdl = true;
                            }
                            else
                            {
                                PBBsMdl = false;
                            }
                        }
                        else
                        {
                            PBBsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBs"));
                                }

                                if (PBBsResult && PBBsMethod && !PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBs"));
                                }

                                if (!PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBs"));
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion
                    #region PPBEs
                    else if (verificationItemCount >= 14 && verificationItemCount <= 23)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBEsTested = false;
                        }
                        else
                        {
                            PBBEsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBEsResult = true;
                        }
                        else
                        {
                            PBBEsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBEsMethod = true;
                        }
                        else
                        {
                            PBBEsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBEsMdl = true;
                            }
                            else
                            {
                                PBBEsMdl = false;
                            }
                        }
                        else
                        {
                            PBBEsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBEs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBEs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBEs"));
                                }

                                if (PBBEsResult && PBBEsMethod && !PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBEs"));
                                }

                                if (!PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBEs"));
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Cl)(Br)
                    else if (verificationItemCount == 28 || verificationItemCount == 29)
                    {
                        string elementName = verificationItemCount == 28 ? "Cl" : "Br";

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Solder paste")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal == s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }

                        clPlusBrResultVal += vertificationResultVal;


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成                       
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (F)
                    else if (verificationItemCount == 30)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"F: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "F"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        vertificationResult = true;
                        /*F的檢測結果不論數值為何，只要有檢測到就透過outllook發郵件通知相關人員*/
                        if (vertificationResultVal != 0 && autoVertificationFlag == true)
                        {
                            Mail.Outlook_to_F(supplierNumer, homogeneousNumber, row.Result);
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "F"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "F"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "F"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PFOS/PFOA
                    else if (verificationItemCount == 31 || verificationItemCount == 32)
                    {
                        /*
                        string elementName = verificationItemCount == 31 ? "PFOS" : "PFOA";
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion

                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion

                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (Sb)
                    else if (verificationItemCount == 33)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Sb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Sb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal <= s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion

                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion

                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Sb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Sb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Sb"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (Be)
                    else if (verificationItemCount == 34)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Be: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Be"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定   
                        if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal <= s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Beryllium copper")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(Be)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }

                        #endregion
                        
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Be"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Be"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Be"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (PVC)                    
                    else if (verificationItemCount == 35)
                    {
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PVC: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PVC"));
                            }
                            verificationItemCount++;
                            continue;

                        }

                        #region 檢測結果判定    
                        vertificationResult = true;
                        #endregion
                       
                        #region 檢測方法判定
                        vertificationMethod = true;
                        #endregion
                       
                        #region mdl標準檢測
                        vertificationMethod = true;
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PVC"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PVC"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PVC"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region 其他有規則
                    else
                    {
                        if (verificationItemCount == 36)
                        {
                            break;
                        }

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", ((withRuleItems)verificationItemCount).ToString()));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", ((withRuleItems)verificationItemCount).ToString()));
                            }
                        }
                        #endregion
                    }
                    #endregion

                    verificationItemCount++;
                }
            }
            #endregion
            #region 3b
            else if (VerificationScenarioIndex == 4)
            {
                if (reportParaments != null)
                {
                    resultList.Add($"報告號碼: {reportParaments.reportNunber}\r\n");
                    resultList.Add($"廠商名稱: {reportParaments.company}\r\n");
                    resultList.Add($"測試期間的最後日期: {reportParaments.finalExtractedDate}\r\n");
                    string status = DateTime.Parse(reportParaments.finalExtractedDate).AddYears(1) > DateTime.Now ? "有效" : "過期";
                    resultList.Add($"效期: {status}\r\n");
                    resultList.Add($"Sample Name: {reportParaments.sampleName}\r\n");
                    resultList.Add($"Style Item No.: {reportParaments.styleItemNo}\r\n");
                    resultList.Add($"樣品描述: {reportParaments.testParts}\r\n\r\n");
                }

                //所有檢查項目列表
                string[] specItems = new[] {
                    "Lead", "Cadmium", "Mercury", "Hexavalent Chromium",
                    "Monobromobiphenyl","Dibromobiphenyl","Tribromobiphenyl","Tetrabromobiphenyl","Pentabromobiphenyl","Hexabromobiphenyl","Heptabromobiphenyl","Octabromobiphenyl","Nonabromobiphenyl","Decabromobiphenyl",
                    "Monobromodiphenyl ether","Dibromodiphenyl ether","Tribromodiphenyl ether","Tetrabromodiphenyl ether","Pentabromodiphenyl ether","Hexabromodiphenyl ether","Heptabromodiphenyl ether","Octabromodiphenyl ether","Nonabromodiphenyl ether","Decabromodiphenyl ether",
                    "DIBP","DEHP","BBP","DBP",
                    "Chlorine","Bromine","Fluorine","PFOS",
                    "PFOA","Antimony","Beryllium","PVC",
                    "PFAS"
                };




                foreach (var s in specsStandards)
                {
                    bool vertificationMethod = false;
                    bool vertificationResult = false;
                    bool vertificationMdl = false;

                    float vertificationResultVal = 0;
                    float vertificationMdlVal = 0;


                    //沒有檢測的直接跳過並記錄在報告裡面
                    var row = manager.reportRowData.FirstOrDefault(r => r.TestItem.Contains(specItems[verificationItemCount]));

                    /*---------------------------------------------------------------------*/
                    #region 檢測結果數值轉換
                    if (row.Result.ToLower() == "n.d.")
                    {
                        vertificationResultVal = 0;
                    }
                    else if (float.TryParse(row.Result, out float parsedVal))
                    {
                        vertificationResultVal = parsedVal;
                    }
                    #endregion
                    /*---------------------------------------------------------------------*/
                    #region (Pb) 編號0
                    if (verificationItemCount == 0)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Pb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Pb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "wire" || homogeneousMaterialType == "lead frame")
                        {
                            if (vertificationResultVal < s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal < s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(pb)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (tpiCompany == 5)
                            {
                                if (vertificationMdlVal < 5)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Pb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Pb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Pb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Pb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Pb"));
                            }
                        }
                        #endregion


                    }
                    #endregion
                    #region Cr(VI) 編號3
                    else if (verificationItemCount == 3)
                    {
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Cr(VI): 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Cr(VI)"));
                            }
                            verificationItemCount++;
                            continue;
                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (homogeneousMaterialType == "Immersion Tin" || homogeneousMaterialType == "Plating/Coatings layer" ||
                           homogeneousMaterialType == "Solder paste")
                        {
                            if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (row.Method.Contains(s.testMethod1))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1)
                        {
                            if (row.Method.Contains(s.testMethod2))
                            {
                                vertificationMethod = true;
                            }
                            else
                            {
                                vertificationMethod = false;
                            }
                        }


                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (MetalNonmetalIndex == 0)
                            {
                                if (vertificationMdlVal == 0.1 || vertificationMdlVal < 0.1)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else if (MetalNonmetalIndex == 1)
                            {
                                if (vertificationMdlVal <= 8)
                                {
                                    vertificationMdl = true;
                                }
                                else
                                {
                                    vertificationMdl = false;
                                }
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Cr(VI): result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Cr(VI): " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Cr(VI)"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Cr(VI)"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Cr(VI)"));
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region PPBs
                    else if (verificationItemCount >= 4 && verificationItemCount <= 13)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBsTested = false;
                        }
                        else
                        {
                            PBBsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBsResult = true;
                        }
                        else
                        {
                            PBBsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBsMethod = true;
                        }
                        else
                        {
                            PBBsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBsMdl = true;
                            }
                            else
                            {
                                PBBsMdl = false;
                            }
                        }
                        else
                        {
                            PBBsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBsTested && verificationItemCount == 13)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add($"PPBs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBs: " + generateVerificationReport(PBBsMethod, PBBsResult, PBBsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBsResult && !PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBs"));
                                }

                                if (PBBsResult && PBBsMethod && !PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBs"));
                                }

                                if (!PBBsResult && PBBsMethod && PBBsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBs"));
                                }
                            }
                        }
                        #endregion

                    }
                    #endregion
                    #region PPBEs
                    else if (verificationItemCount >= 14 && verificationItemCount <= 23)
                    {
                        if ((row == null || row.Result == "---"))
                        {
                            PBBEsTested = false;
                        }
                        else
                        {
                            PBBEsTested = true;
                        }
                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            PBBEsResult = true;
                        }
                        else
                        {
                            PBBEsResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            PBBEsMethod = true;
                        }
                        else
                        {
                            PBBEsMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                PBBEsMdl = true;
                            }
                            else
                            {
                                PBBEsMdl = false;
                            }
                        }
                        else
                        {
                            PBBEsMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PPBEs: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PPBEs"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        else if (PBBEsTested && verificationItemCount == 23)
                        {
                            if (!autoVertificationFlag)
                            {
                                if (PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                                }
                                else if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add($"PPBEs: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                                else
                                {
                                    resultList.Add($"PPBEs: " + generateVerificationReport(PBBEsMethod, PBBEsResult, PBBEsMdl) + "\r\n");
                                }
                            }
                            else
                            {
                                if (PBBEsResult && !PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PPBEs"));
                                }

                                if (PBBEsResult && PBBEsMethod && !PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PPBEs"));
                                }

                                if (!PBBEsResult && PBBEsMethod && PBBEsMdl)
                                {
                                    resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PPBEs"));
                                }
                            }
                        }
                        #endregion
                    }
                    #endregion
                    #region (Cl)(Br)
                    else if (verificationItemCount == 28 || verificationItemCount == 29)
                    {
                        /*
                        string elementName = verificationItemCount == 28 ? "Cl" : "Br";

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Solder paste")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal == s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }

                        clPlusBrResultVal += vertificationResultVal;


                        #endregion
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成                       
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }

                            if (clPlusBrResultVal >= 1000 && verificationItemCount == 29)
                            {
                                resultList.Add($"Cl + Br: 檢測值超標  Exceed SPIL’s limitation\r\n");
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (F)
                    else if (verificationItemCount == 30)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"F: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "F"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定
                        vertificationResult = true;
                        
                        if (vertificationResultVal != 0 && autoVertificationFlag == true)
                        {
                            Mail.Outlook_to_F(supplierNumer, homogeneousNumber, row.Result);
                        }
                        #endregion
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion

                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"F: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"F: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "F"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "F"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "F"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region PFOS/PFOA
                    else if (verificationItemCount == 31 || verificationItemCount == 32)
                    {
                        /*
                        string elementName = verificationItemCount == 31 ? "PFOS" : "PFOA";
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{elementName}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", elementName));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion

                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion

                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{elementName}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{elementName}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }

                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", elementName));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", elementName));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", elementName));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (Sb)
                    else if (verificationItemCount == 33)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Sb: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Sb"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal <= s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion

                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion

                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Sb: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Sb: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Sb"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Sb"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Sb"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (Be)
                    else if (verificationItemCount == 34)
                    {
                        /*
                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"Be: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "Be"));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定   
                        if (MetalNonmetalIndex == 0)
                        {
                            if (vertificationResultVal <= s.Metal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else if (MetalNonmetalIndex == 1 || homogeneousMaterialType == "Beryllium copper")
                        {
                            if (vertificationResultVal <= s.NonMetal)
                            {
                                vertificationResult = true;
                            }
                            else
                            {
                                vertificationResult = false;
                            }
                        }
                        else
                        {
                            throw new Exception("(Be)金屬非金屬參數錯誤 MetalNonmetalIndex is invalid");
                        }

                        #endregion
                        
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1) || row.Method.Contains(s.testMethod2))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"Be: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"Be: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "Be"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "Be"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "Be"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region (PVC)                    
                    else if (verificationItemCount == 35)
                    {
                        /*
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"PVC: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", "PVC"));
                            }
                            verificationItemCount++;
                            continue;

                        }

                        #region 檢測結果判定    
                        vertificationResult = true;
                        #endregion
                       
                        #region 檢測方法判定
                        vertificationMethod = true;
                        #endregion
                       
                        #region mdl標準檢測
                        vertificationMethod = true;
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"PVC: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"PVC: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", "PVC"));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", "PVC"));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", "PVC"));
                            }
                        }
                        #endregion
                        */
                    }
                    #endregion
                    #region 其他有規則
                    else
                    {
                        if (verificationItemCount == 36)
                        {
                            break;
                        }

                        #region 未檢測
                        if (row.Result == "---")
                        {
                            if (!autoVertificationFlag)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: 未檢測 Test item is not found\r\n");
                            }
                            else
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-06", ((withRuleItems)verificationItemCount).ToString()));
                            }
                            verificationItemCount++;
                            continue;

                        }
                        #endregion

                        #region 檢測結果判定                        
                        if (vertificationResultVal == s.Metal)
                        {
                            vertificationResult = true;
                        }
                        else
                        {
                            vertificationResult = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region 檢測方法判定
                        if (row.Method.Contains(s.testMethod1))
                        {
                            vertificationMethod = true;
                        }
                        else
                        {
                            vertificationMethod = false;
                        }
                        #endregion
                        /*---------------------------------------------------------------------*/
                        #region mdl標準檢測
                        if (float.TryParse(row.MDL, out float parsedMdl))
                        {
                            vertificationMdlVal = parsedMdl;
                            if (vertificationMdlVal <= s.MdlValue)
                            {
                                vertificationMdl = true;
                            }
                            else
                            {
                                vertificationMdl = false;
                            }
                        }
                        else
                        {
                            vertificationMdl = false;
                        }
                        #endregion

                        #region 報告生成
                        if (!autoVertificationFlag)
                        {
                            if (vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + $" result: {row.Result}" + $" MDL: {row.MDL} \r\n");

                            }
                            else if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: result: {row.Result}  MDL: {row.MDL}  " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                            else
                            {
                                resultList.Add($"{((withRuleItems)verificationItemCount).ToString()}: " + generateVerificationReport(vertificationMethod, vertificationResult, vertificationMdl) + "\r\n");
                            }
                        }
                        else
                        {
                            if (vertificationResult && !vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-07", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (vertificationResult && vertificationMethod && !vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-08", ((withRuleItems)verificationItemCount).ToString()));
                            }

                            if (!vertificationResult && vertificationMethod && vertificationMdl)
                            {
                                resultList.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-09", ((withRuleItems)verificationItemCount).ToString()));
                            }
                        }
                        #endregion
                    }
                    #endregion

                    verificationItemCount++;
                }
            }
            #endregion

            return resultList;
        }

        public string generateVerificationReport(bool vertificationMethod, bool vertificationResult, bool vertificationMdl)
        { 
            if(vertificationMethod && vertificationResult && vertificationMdl)
            {
                return "PASS";
            }
            else if (!vertificationResult)
            {
                return "檢測值超標 Exceed SPIL’s limitation";
            }
            else if (!vertificationMdl)
            {
                return "超出矽品認可MDL Not qualified MDL";
            }
            else if (!vertificationMethod && vertificationResult && vertificationMdl)
            {
                return "檢測方法未經矽品認可 Not qualified test method";
            }
            else
            {
                return "Something Wrong~~";
            }
        }
    }
}
