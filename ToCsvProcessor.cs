﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.XPath;
using System.Diagnostics;

namespace TabulateSmarterTestResults
{

    class ToCsvProcessor : ITestResultProcessor
    {
        static string[] sStudentFieldNames = new string[]
        {
            "StateAbbreviation",
            "ResponsibleDistrictIdentifier",
            "OrganizationName",
            "ResponsibleSchoolIdentifier",
            "NameOfInstitution",
            "StudentIdentifier",
            "ExternalSSID",
            "FirstName",
            "MiddleName",
            "LastOrSurname",
            "Sex",
            "Birthdate",
            "GradeLevelWhenAssessed",
            "HispanicOrLatinoEthnicity",
            "AmericanIndianOrAlaskaNative",
            "Asian",
            "BlackOrAfricanAmerican",
            "NativeHawaiianOrOtherPacificIslander",
            "White",
            "DemographicRaceTwoOrMoreRaces",
            "IDEAIndicator",
            "LEPStatus",
            "Section504Status",
            "EconomicDisadvantageStatus",
            "MigrantStatus",
            "Group1Id",
            "Group1Text",
            "Group2Id",
            "Group2Text",
            "Group3Id",
            "Group3Text",
            "Group4Id",
            "Group4Text",
            "Group5Id",
            "Group5Text",
            "Group6Id",
            "Group6Text",
            "Group7Id",
            "Group7Text",
            "Group8Id",
            "Group8Text",
            "Group9Id",
            "Group9Text",
            "Group10Id",
            "Group10Text",
            "AssessmentGuid",
            "AssessmentSessionLocationId",
            "AssessmentSessionLocation",
            "AssessmentAdministrationFinishDate",
            "AssessmentYear",
            "AssessmentType",
            "AssessmentAcademicSubject",
            "AssessmentLevelForWhichDesigned",
            "AssessmentSubtestResultScoreValue",
            "AssessmentSubtestMinimumValue",
            "AssessmentSubtestMaximumValue",
            "AssessmentPerformanceLevelIdentifier",
            "AssessmentSubtestResultScoreClaim1Value",
            "AssessmentSubtestClaim1MinimumValue",
            "AssessmentSubtestClaim1MaximumValue",
            "AssessmentClaim1PerformanceLevelIdentifier",
            "AssessmentSubtestResultScoreClaim2Value",
            "AssessmentSubtestClaim2MinimumValue",
            "AssessmentSubtestClaim2MaximumValue",
            "AssessmentClaim2PerformanceLevelIdentifier",
            "AssessmentSubtestResultScoreClaim3Value",
            "AssessmentSubtestClaim3MinimumValue",
            "AssessmentSubtestClaim3MaximumValue",
            "AssessmentClaim3PerformanceLevelIdentifier",
            "AssessmentSubtestResultScoreClaim4Value",
            "AssessmentSubtestClaim4MinimumValue",
            "AssessmentSubtestClaim4MaximumValue",
            "AssessmentClaim4PerformanceLevelIdentifier",
            "AccommodationAmericanSignLanguage",
            "AccommodationBraille",
            "AccommodationClosedCaptioning",
            "AccommodationTextToSpeech",
            "AccommodationAbacus",
            "AccommodationAlternateResponseOptions",
            "AccommodationCalculator",
            "AccommodationMultiplicationTable",
            "AccommodationPrintOnDemand",
            "AccommodationPrintOnDemandItems",
            "AccommodationReadAloud",
            "AccommodationScribe",
            "AccommodationSpeechToText",
            "AccommodationStreamlineMode",
            "AccommodationNoiseBuffer"
        };

        static string[] sItemFieldNames = new string[]
        {
            "key", // ItemId
            "studentId", // May be StudentIdentifier or AlternateSSID
            "segmentId",
            "position",
            "clientId",
            "operational",
            "isSelected",
            "format",
            "score",
            "scoreStatus",
            "adminDate",
            "numberVisits",
            "strand",
            "contentLevel",
            "pageNumber",
            "pageVisits",
            "pageTime",
            "dropped"
        };

        static XPathExpression sXp_StateAbbreviation = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='StateAbbreviation' and @context='FINAL']/@value");
        static XPathExpression sXp_DistrictId = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='ResponsibleDistrictIdentifier' and @context='FINAL']/@value");
        static XPathExpression sXp_DistrictName = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='OrganizationName' and @context='FINAL']/@value");
        static XPathExpression sXp_SchoolId = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='ResponsibleInstitutionIdentifier' and @context='FINAL']/@value");
        static XPathExpression sXp_SchoolName = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='NameOfInstitution' and @context='FINAL']/@value");
        static XPathExpression sXp_StudentIdentifier = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='StudentIdentifier' and @context='FINAL']/@value");
        static XPathExpression sXp_AlternateSSID = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='AlternateSSID' and @context='FINAL']/@value");
        static XPathExpression sXp_FirstName = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='FirstName' and @context='FINAL']/@value");
        static XPathExpression sXp_MiddleName = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='MiddleName' and @context='FINAL']/@value");
        static XPathExpression sXp_LastOrSurname = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='LastOrSurname' and @context='FINAL']/@value");
        static XPathExpression sXp_Sex = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='Sex' and @context='FINAL']/@value");
        static XPathExpression sXp_Birthdate = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='Birthdate' and @context='FINAL']/@value");
        static XPathExpression sXp_GradeLevelWhenAssessed = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='GradeLevelWhenAssessed' and @context='FINAL']/@value");
        static XPathExpression sXp_HispanicOrLatinoEthnicity = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='HispanicOrLatinoEthnicity' and @context='FINAL']/@value");
        static XPathExpression sXp_AmericanIndianOrAlaskaNative = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='AmericanIndianOrAlaskaNative' and @context='FINAL']/@value");
        static XPathExpression sXp_Asian = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='Asian' and @context='FINAL']/@value");
        static XPathExpression sXp_BlackOrAfricanAmerican = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='BlackOrAfricanAmerican' and @context='FINAL']/@value");
        static XPathExpression sXp_NativeHawaiianOrOtherPacificIslander = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='NativeHawaiianOrOtherPacificIslander' and @context='FINAL']/@value");
        static XPathExpression sXp_White = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='White' and @context='FINAL']/@value");
        static XPathExpression sXp_DemographicRaceTwoOrMoreRaces = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='DemographicRaceTwoOrMoreRaces' and @context='FINAL']/@value");
        static XPathExpression sXp_IDEAIndicator = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='IDEAIndicator' and @context='FINAL']/@value");
        static XPathExpression sXp_LEPStatus = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='LEPStatus' and @context='FINAL']/@value");
        static XPathExpression sXp_Section504Status = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='Section504Status' and @context='FINAL']/@value");
        static XPathExpression sXp_EconomicDisadvantageStatus = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='EconomicDisadvantageStatus' and @context='FINAL']/@value");
        static XPathExpression sXp_MigrantStatus = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='MigrantStatus' and @context='FINAL']/@value");
        static XPathExpression sXp_StudentGroupName = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='StudentGroupName' and @context='FINAL']/@value");
        static XPathExpression sXp_AssessmentId = XPathExpression.Compile("/TDSReport/Test/@testId");
        static XPathExpression sXp_TestSessionId = XPathExpression.Compile("/TDSReport/Opportunity/@sessionId");
        static XPathExpression sXp_SchoolYear = XPathExpression.Compile("/TDSReport/Test/@academicYear");
        static XPathExpression sXp_AssessmentType = XPathExpression.Compile("/TDSReport/Test/@assessmentType");
        static XPathExpression sXp_Subject = XPathExpression.Compile("/TDSReport/Test/@subject");
        static XPathExpression sXp_TestGrade = XPathExpression.Compile("/TDSReport/Test/@grade");
        static XPathExpression sXp_ScaleScore = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and @measureLabel='ScaleScore']/@value");
        static XPathExpression sXp_ScaleScoreStandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and @measureLabel='ScaleScore']/@standardError");
        static XPathExpression sXp_ScaleScoreAchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and @measureLabel='PerformanceLevel']/@value");
        // Claim 1 for ELA is labeled "SOCK_R" for math it's labeled "1"
        static XPathExpression sXp_ClaimScore1 = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_R' or @measureOf='1') and @measureLabel='ScaleScore']/@value");
        static XPathExpression sXp_ClaimScore1StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_R' or @measureOf='1') and @measureLabel='ScaleScore']/@standardError");
        static XPathExpression sXp_ClaimScore1AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_R' or @measureOf='1') and @measureLabel='PerformanceLevel']/@value");
        // Claim 2 for ELA is labeled "2-W" for math it's labeled "SOCK_2"
        static XPathExpression sXp_ClaimScore2 = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='2-W' or @measureOf='SOCK_2') and @measureLabel='ScaleScore']/@value");
        static XPathExpression sXp_ClaimScore2StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='2-W' or @measureOf='SOCK_2') and @measureLabel='ScaleScore']/@standardError");
        static XPathExpression sXp_ClaimScore2AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='2-W' or @measureOf='SOCK_2') and @measureLabel='PerformanceLevel']/@value");
        // Claim 3 for ELA is labeled "SOCK_LS" for math it's labeled "3"
        static XPathExpression sXp_ClaimScore3 = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_LS' or @measureOf='3') and @measureLabel='ScaleScore']/@value");
        static XPathExpression sXp_ClaimScore3StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_LS' or @measureOf='3') and @measureLabel='ScaleScore']/@standardError");
        static XPathExpression sXp_ClaimScore3AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_LS' or @measureOf='3') and @measureLabel='PerformanceLevel']/@value");
        // Claim 4 for ELA is labeled "4-CR" for math it doesn't exist
        static XPathExpression sXp_ClaimScore4 = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='4-CR' and @measureLabel='ScaleScore']/@value");
        static XPathExpression sXp_ClaimScore4StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='4-CR' and @measureLabel='ScaleScore']/@standardError");
        static XPathExpression sXp_ClaimScore4AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='4-CR' and @measureLabel='PerformanceLevel']/@value");
        // Matches all accessibility codes
        static XPathExpression sXP_AccessibilityCodes = XPathExpression.Compile("/TDSReport/Opportunity/Accommodation/@code");

        // Item Level Data
        static XPathExpression sXP_Item = XPathExpression.Compile("/TDSReport/Opportunity/Item");
        static XPathExpression sXp_ItemKey = XPathExpression.Compile("@key");
        static XPathExpression sXp_ItemBankKey = XPathExpression.Compile("@bankKey");
        static XPathExpression sXp_SegmentId = XPathExpression.Compile("@segmentId");
        static XPathExpression sXp_ItemPosition = XPathExpression.Compile("@position");
        static XPathExpression sXp_ClientId = XPathExpression.Compile("@clientId");
        static XPathExpression sXp_Operational = XPathExpression.Compile("@operational");
        static XPathExpression sXp_IsSelected = XPathExpression.Compile("@isSelected");
        static XPathExpression sXp_ItemType = XPathExpression.Compile("@format");
        static XPathExpression sXp_ItemScore = XPathExpression.Compile("@score");
        static XPathExpression sXp_ScoreStatus = XPathExpression.Compile("@scoreStatus");
        static XPathExpression sXp_AdminDate = XPathExpression.Compile("@adminDate");
        static XPathExpression sXp_NumberVisits = XPathExpression.Compile("@numberVisits");
        static XPathExpression sXp_Strand = XPathExpression.Compile("@strand");
        static XPathExpression sXp_ContentLevel = XPathExpression.Compile("@contentLevel");
        static XPathExpression sXp_PageNumber = XPathExpression.Compile("@pageNumber");
        static XPathExpression sXp_PageVisits = XPathExpression.Compile("@pageVisits");
        static XPathExpression sXp_PageTime = XPathExpression.Compile("@pageTime");
        static XPathExpression sXp_Dropped = XPathExpression.Compile("@dropped");

        static Dictionary<string, int> sAccessibilityCodeMapping;

        static ToCsvProcessor()
        {
            sAccessibilityCodeMapping = new Dictionary<string, int>();
            sAccessibilityCodeMapping.Add("TDS_ASL0", 73); // American Sign Language
            sAccessibilityCodeMapping.Add("TDS_ASL1", 73);
            sAccessibilityCodeMapping.Add("TDS_ClosedCap0", 75); // Closed Captioning
            sAccessibilityCodeMapping.Add("TDS_ClosedCap1", 75);
            sAccessibilityCodeMapping.Add("ENU-Braille", 74); // Braille
            sAccessibilityCodeMapping.Add("TDS_PoD_Stim", 81); // Print on Demand Stimuli
            sAccessibilityCodeMapping.Add("TDS_PoD_Item", 82); // Print on Demand Item
            sAccessibilityCodeMapping.Add("TDS_TS_Accessibility", 86); // Streamline
            sAccessibilityCodeMapping.Add("TDS_TTS0", 76); // Text to Speech
            sAccessibilityCodeMapping.Add("TDS_TTS_Item", 76);
            sAccessibilityCodeMapping.Add("TDS_TTS_Stim", 76);
            sAccessibilityCodeMapping.Add("TDS_TTS_Stim&TDS_TTS_Item", 76);
            sAccessibilityCodeMapping.Add("NEA_AR", 78); // Non-Embedded Alternate Response Options
            sAccessibilityCodeMapping.Add("NEA_RA_Stimuli", 83); // Non-Embedded Read Aloud
            sAccessibilityCodeMapping.Add("NEA_SC_WritItems", 84); // Non-Embedded Scribe
            sAccessibilityCodeMapping.Add("NEA_STT", 85); // Non-Embedded Speech to Text
            sAccessibilityCodeMapping.Add("NEA_Abacus", 77); // Non-Embedded Abacus
            sAccessibilityCodeMapping.Add("NEA_Calc", 79); // Non-Embedded Calculator
            sAccessibilityCodeMapping.Add("NEA_MT", 80); // Non-Embedded Multiplication Table
            sAccessibilityCodeMapping.Add("NEA_NoiseBuf", 87); // Non-Embedded Noise Buffer
        }
      
        Parse.CsvWriter mSWriter;
        Parse.CsvWriter mIWriter;

        public ToCsvProcessor(string osFilename, string oiFilename)
        {
#if !DEBUG
            if (File.Exists(osFilename)) throw new ApplicationException(string.Format("Output file, '{0}' already exists.", osFilename));
            if (File.Exists(oiFilename)) throw new ApplicationException(string.Format("Output file, '{0}' already exists.", oiFilename));
#endif
            if (osFilename != null)
            {
                mSWriter = new Parse.CsvWriter(osFilename, false);
                mSWriter.Write(sStudentFieldNames);
            }
            if (oiFilename != null)
            {
                mIWriter = new Parse.CsvWriter(oiFilename, false);
                mIWriter.Write(sItemFieldNames);
            }
        }

        public void ProcessResult(Stream input)
        {
            XPathDocument doc = new XPathDocument(input);
            XPathNavigator nav = doc.CreateNavigator();

            // Retrieve the student fields
            string[] studentFields = new string[sStudentFieldNames.Length];

            studentFields[0] = nav.Eval(sXp_StateAbbreviation);
            studentFields[1] = nav.Eval(sXp_DistrictId);
            studentFields[2] = nav.Eval(sXp_DistrictName);
            studentFields[3] = nav.Eval(sXp_SchoolId);
            studentFields[4] = nav.Eval(sXp_SchoolName);
            string studentId = nav.Eval(sXp_StudentIdentifier);
            studentFields[5] = studentId;
            studentFields[6] = nav.Eval(sXp_AlternateSSID);
            studentFields[7] = nav.Eval(sXp_FirstName);
            studentFields[8] = nav.Eval(sXp_MiddleName);
            studentFields[9] = nav.Eval(sXp_LastOrSurname);
            studentFields[10] = nav.Eval(sXp_Sex);
            studentFields[11] = nav.Eval(sXp_Birthdate);
            studentFields[12] = nav.Eval(sXp_GradeLevelWhenAssessed);
            studentFields[13] = nav.Eval(sXp_HispanicOrLatinoEthnicity);
            studentFields[14] = nav.Eval(sXp_AmericanIndianOrAlaskaNative);
            studentFields[15] = nav.Eval(sXp_Asian);
            studentFields[16] = nav.Eval(sXp_BlackOrAfricanAmerican);
            studentFields[17] = nav.Eval(sXp_NativeHawaiianOrOtherPacificIslander);
            studentFields[18] = nav.Eval(sXp_White);
            studentFields[19] = nav.Eval(sXp_DemographicRaceTwoOrMoreRaces);
            studentFields[20] = nav.Eval(sXp_IDEAIndicator);
            studentFields[21] = nav.Eval(sXp_LEPStatus);
            studentFields[22] = nav.Eval(sXp_Section504Status);
            studentFields[23] = nav.Eval(sXp_EconomicDisadvantageStatus);
            studentFields[24] = nav.Eval(sXp_MigrantStatus);

            // Up to 10 student groups
            {
                XPathNodeIterator nodes = nav.Select(sXp_StudentGroupName);
                {
                    int i = 0;
                    while (i < 10 && nodes.MoveNext())
                    {
                        string value = nodes.Current.ToString();
                        studentFields[i * 2 + 25] = value;
                        studentFields[i * 2 + 26] = value;
                        ++i;
                    }
                    while (i < 10)
                    {
                        studentFields[i * 2 + 25] = string.Empty;
                        studentFields[i * 2 + 26] = string.Empty;
                        ++i;
                    }
                }
            }

            studentFields[45] = nav.Eval(sXp_AssessmentId);
            studentFields[46] = nav.Eval(sXp_TestSessionId);
            studentFields[47] = string.Empty; // AssessmentLocation
            studentFields[48] = string.Empty; // AssessmentAdministrationFinishDate
            studentFields[49] = nav.Eval(sXp_SchoolYear);
            studentFields[50] = nav.Eval(sXp_AssessmentType);
            studentFields[51] = nav.Eval(sXp_Subject);
            studentFields[52] = nav.Eval(sXp_TestGrade);
            ProcessScores(nav, sXp_ScaleScore, sXp_ScaleScoreStandardError, sXp_ScaleScoreAchievementLevel, studentFields, 53);
            ProcessScores(nav, sXp_ClaimScore1, sXp_ClaimScore1StandardError, sXp_ClaimScore1AchievementLevel, studentFields, 57);
            ProcessScores(nav, sXp_ClaimScore2, sXp_ClaimScore2StandardError, sXp_ClaimScore2AchievementLevel, studentFields, 61);
            ProcessScores(nav, sXp_ClaimScore3, sXp_ClaimScore3StandardError, sXp_ClaimScore3AchievementLevel, studentFields, 65);
            ProcessScores(nav, sXp_ClaimScore4, sXp_ClaimScore4StandardError, sXp_ClaimScore4AchievementLevel, studentFields, 69);

            // Preload accommodation fields with empty string
            for (int i = 73; i < sStudentFieldNames.Length; ++i) studentFields[i] = string.Empty;

            // Process accommodations
            {
                XPathNodeIterator nodes = nav.Select(sXP_AccessibilityCodes);
                {
                    while (nodes.MoveNext())
                    {
                        string code = nodes.Current.ToString();
                        int fieldIndex;
                        if (sAccessibilityCodeMapping.TryGetValue(code, out fieldIndex))
                        {
                            studentFields[fieldIndex] = code;
                        }
                    }
                }
            }

            // Write one line to the CSV
            if (mSWriter != null)
                mSWriter.Write(studentFields);

            // Report item data
            {
                XPathNodeIterator nodes = nav.Select(sXP_Item);
                while (nodes.MoveNext())
                {
                    // Collect the item fields
                    string[] itemFields = new string[sItemFieldNames.Length];

                    XPathNavigator node = nodes.Current;

                    itemFields[0] = string.Concat(node.Eval(sXp_ItemBankKey), "-", node.Eval(sXp_ItemKey));
                    itemFields[1] = studentId;
                    itemFields[2] = node.Eval(sXp_SegmentId);
                    itemFields[3] = node.Eval(sXp_ItemPosition);
                    itemFields[4] = node.Eval(sXp_ClientId);
                    itemFields[5] = node.Eval(sXp_Operational);
                    itemFields[6] = node.Eval(sXp_IsSelected);
                    itemFields[7] = node.Eval(sXp_ItemType);
                    itemFields[8] = node.Eval(sXp_ItemScore);
                    itemFields[9] = node.Eval(sXp_ScoreStatus);
                    itemFields[10] = node.Eval(sXp_AdminDate);
                    itemFields[11] = node.Eval(sXp_NumberVisits);
                    itemFields[12] = node.Eval(sXp_Strand);
                    itemFields[13] = node.Eval(sXp_ContentLevel);
                    itemFields[14] = node.Eval(sXp_PageNumber);
                    itemFields[15] = node.Eval(sXp_PageVisits);
                    itemFields[16] = node.Eval(sXp_PageTime);
                    itemFields[17] = node.Eval(sXp_Dropped);

                    // Write one line to the CSV
                    if (mIWriter != null)
                        mIWriter.Write(itemFields);
                }
            }

        }

        static void ProcessScores(XPathNavigator nav, XPathExpression xp_ScaleScore, XPathExpression xp_StdErr,
            XPathExpression xp_PerfLvl, string[] fields, int index)
        {
            string scaleScore = nav.Eval(xp_ScaleScore);
            string stdErr = nav.Eval(xp_StdErr);
            string perfLvl = nav.Eval(xp_PerfLvl);

            fields[index] = scaleScore;
            fields[index + 3] = perfLvl;

            double scaleScoreF;
            double stdErrF;
            if (double.TryParse(scaleScore, out scaleScoreF) && double.TryParse(stdErr, out stdErrF))
            {
                // MinimumValue
                fields[index + 1] = (scaleScoreF - stdErrF).ToString("F3", System.Globalization.CultureInfo.InvariantCulture);
                // MaximumValue
                fields[index + 2] = (scaleScoreF + stdErrF).ToString("F3", System.Globalization.CultureInfo.InvariantCulture);
            }
            else
            {
                fields[index + 1] = string.Empty;
                fields[index + 2] = string.Empty;
            }
        }

        ~ToCsvProcessor()
        {
            Dispose(false);
        }


        public void Dispose()
        {
            Dispose(true);
        }

        void Dispose(bool disposing)
        {
            if (mSWriter != null)
            {
#if DEBUG
                if (!disposing) Debug.Fail("Failed to dispose CsvWriter");
#endif
                mSWriter.Dispose();
                mSWriter = null;
            }
            if (mIWriter != null)
            {
                mIWriter.Dispose();
                mIWriter = null;
            }
            if (disposing)
            {
                GC.SuppressFinalize(this);
            }
        }
    }

    static class XPathNavitagorHelper
    {
        public static string Eval(this XPathNavigator nav, XPathExpression expression)
        {
            if (expression.ReturnType == XPathResultType.NodeSet)
            {
                XPathNodeIterator nodes = nav.Select(expression);
                if (nodes.MoveNext())
                {
                    return nodes.Current.ToString();
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return nav.Evaluate(expression).ToString();
            }
        }
    }
}
