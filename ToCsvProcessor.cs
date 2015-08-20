using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.XPath;
using System.Diagnostics;

namespace TabulateSmarterTestResults
{

    class XmlCsvMapping
    {
        string mCsvFieldName;
        XPathExpression mXpathQuery;

        public XmlCsvMapping(string csvFieldName, string xpathQuery)
        {
            mCsvFieldName = csvFieldName;
            mXpathQuery = XPathExpression.Compile(xpathQuery);
        }

        public string CsvFieldName
        {
            get { return mCsvFieldName; }
        }

        public XPathExpression XPathQuery
        {
            get { return mXpathQuery; }
        }

    }

    class ToCsvProcessor : ITestResultProcessor
    {
        static string[] sFieldNames = new string[]
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

        static XPathExpression sXp_StateAbbreviation = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='StateAbbreviation' and @context='FINAL']/@value");
        static XPathExpression sXp_ResponsibleDistrictIdentifier = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='ResponsibleDistrictIdentifier' and @context='FINAL']/@value");
        static XPathExpression sXp_OrganizationName = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='OrganizationName' and @context='FINAL']/@value");
        static XPathExpression sXp_ResponsibleSchoolIdentifier = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='ResponsibleInstitutionIdentifier' and @context='FINAL']/@value");
        static XPathExpression sXp_NameOfInstitution = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='NameOfInstitution' and @context='FINAL']/@value");
        static XPathExpression sXp_StudentIdentifier = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='StudentIdentifier' and @context='FINAL']/@value");
        static XPathExpression sXp_ExternalSSID = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='AlternateSSID' and @context='FINAL']/@value");
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
        static XPathExpression sXp_GroupId = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='StudentGroupName' and @context='FINAL']/@value");
        static XPathExpression sXp_AssessmentGuid = XPathExpression.Compile("/TDSReport/Test/@testId");
        static XPathExpression sXp_AssessmentLocationId = XPathExpression.Compile("/TDSReport/Opportunity/@sessionId");
        static XPathExpression sXp_AssessmentYear = XPathExpression.Compile("/TDSReport/Test/@academicYear");
        static XPathExpression sXp_AssessmentType = XPathExpression.Compile("/TDSReport/Test/@assessmentType");
        static XPathExpression sXp_AssessmentAcademicSubject = XPathExpression.Compile("/TDSReport/Test/@subject");
        static XPathExpression sXp_AssessmentLevelForWhichDesigned = XPathExpression.Compile("/TDSReport/Test/@grade");
        // From here forward, names are based on the Logical Data Model rather than the destination CSV field names
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
      
        Parse.CsvWriter mWriter;

        public ToCsvProcessor(string filename)
        {
#if !DEBUG
            if (File.Exists(filename)) throw new ApplicationException(string.Format("Output file, '{0}' already exists.", filename));
#endif
            mWriter = new Parse.CsvWriter(filename, false);

            // Write out the field names in the first line
            mWriter.Write(sFieldNames);
        }

        public void ProcessResult(Stream input)
        {
            XPathDocument doc = new XPathDocument(input);
            XPathNavigator nav = doc.CreateNavigator();

            // Retrieve the fields
            string[] fields = new string[sFieldNames.Length];

            fields[0] = nav.Eval(sXp_StateAbbreviation);
            fields[1] = nav.Eval(sXp_ResponsibleDistrictIdentifier);
            fields[2] = nav.Eval(sXp_OrganizationName);
            fields[3] = nav.Eval(sXp_ResponsibleSchoolIdentifier);
            fields[4] = nav.Eval(sXp_NameOfInstitution);
            fields[5] = nav.Eval(sXp_StudentIdentifier);
            fields[6] = nav.Eval(sXp_ExternalSSID);
            fields[7] = nav.Eval(sXp_FirstName);
            fields[8] = nav.Eval(sXp_MiddleName);
            fields[9] = nav.Eval(sXp_LastOrSurname);
            fields[10] = nav.Eval(sXp_Sex);
            fields[11] = nav.Eval(sXp_Birthdate);
            fields[12] = nav.Eval(sXp_GradeLevelWhenAssessed);
            fields[13] = nav.Eval(sXp_HispanicOrLatinoEthnicity);
            fields[14] = nav.Eval(sXp_AmericanIndianOrAlaskaNative);
            fields[15] = nav.Eval(sXp_Asian);
            fields[16] = nav.Eval(sXp_BlackOrAfricanAmerican);
            fields[17] = nav.Eval(sXp_NativeHawaiianOrOtherPacificIslander);
            fields[18] = nav.Eval(sXp_White);
            fields[19] = nav.Eval(sXp_DemographicRaceTwoOrMoreRaces);
            fields[20] = nav.Eval(sXp_IDEAIndicator);
            fields[21] = nav.Eval(sXp_LEPStatus);
            fields[22] = nav.Eval(sXp_Section504Status);
            fields[23] = nav.Eval(sXp_EconomicDisadvantageStatus);
            fields[24] = nav.Eval(sXp_MigrantStatus);

            // Up to 10 student groups
            {
                XPathNodeIterator nodes = nav.Select(sXp_GroupId);
                {
                    int i = 0;
                    while (i < 10 && nodes.MoveNext())
                    {
                        string value = nodes.Current.ToString();
                        fields[i * 2 + 25] = value;
                        fields[i * 2 + 26] = value;
                        ++i;
                    }
                    while (i < 10)
                    {
                        fields[i * 2 + 25] = string.Empty;
                        fields[i * 2 + 26] = string.Empty;
                        ++i;
                    }
                }
            }

            fields[45] = nav.Eval(sXp_AssessmentGuid);
            fields[46] = nav.Eval(sXp_AssessmentLocationId);
            fields[47] = string.Empty; // AssessmentLocation
            fields[48] = string.Empty; // AssessmentAdministrationFinishDate
            fields[49] = nav.Eval(sXp_AssessmentYear);
            fields[50] = nav.Eval(sXp_AssessmentType);
            fields[51] = nav.Eval(sXp_AssessmentAcademicSubject);
            fields[52] = nav.Eval(sXp_AssessmentLevelForWhichDesigned);
            ProcessScores(nav, sXp_ScaleScore, sXp_ScaleScoreStandardError, sXp_ScaleScoreAchievementLevel, fields, 53);
            ProcessScores(nav, sXp_ClaimScore1, sXp_ClaimScore1StandardError, sXp_ClaimScore1AchievementLevel, fields, 57);
            ProcessScores(nav, sXp_ClaimScore2, sXp_ClaimScore2StandardError, sXp_ClaimScore2AchievementLevel, fields, 61);
            ProcessScores(nav, sXp_ClaimScore3, sXp_ClaimScore3StandardError, sXp_ClaimScore3AchievementLevel, fields, 65);
            ProcessScores(nav, sXp_ClaimScore4, sXp_ClaimScore4StandardError, sXp_ClaimScore4AchievementLevel, fields, 69);

            // Preload accommodation fields with empty string
            for (int i = 73; i < sFieldNames.Length; ++i) fields[i] = string.Empty;

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
                            fields[fieldIndex] = code;
                        }
                    }
                }
            }

            // Write one line to the CSV
            mWriter.Write(fields);
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
            if (mWriter != null)
            {
#if DEBUG
                if (!disposing) Debug.Fail("Failed to dispose CsvWriter");
#endif
                mWriter.Dispose();
                mWriter = null;
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
