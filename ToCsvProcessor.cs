using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.XPath;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Globalization;

namespace TabulateSmarterTestResults
{

    class ToCsvProcessor : ITestResultProcessor
    {

        // Student Field names for the Data Warehouse format
        enum DwStudentFieldNames : int
        {
            StateAbbreviation, // 0
            ResponsibleDistrictIdentifier, // 1
            OrganizationName, // 2
            ResponsibleSchoolIdentifier, // 3
            NameOfInstitution, // 4
            StudentIdentifier, // 5
            ExternalSSID, // 6
            FirstName, // 7
            MiddleName, // 8
            LastOrSurname, // 9
            Sex, // 10
            Birthdate, // 11
            GradeLevelWhenAssessed, // 12
            HispanicOrLatinoEthnicity, // 13
            AmericanIndianOrAlaskaNative, // 14
            Asian, // 15
            BlackOrAfricanAmerican, // 16
            NativeHawaiianOrOtherPacificIslander, // 17
            White, // 18
            DemographicRaceTwoOrMoreRaces, // 19
            IDEAIndicator, // 20
            LEPStatus, // 21
            Section504Status, // 22
            EconomicDisadvantageStatus, // 23
            MigrantStatus, // 24
            Group1Id, // 25
            Group1Text, // 26
            Group2Id, // 27
            Group2Text, // 28
            Group3Id, // 29
            Group3Text, // 30
            Group4Id, // 31
            Group4Text, // 32
            Group5Id, // 33
            Group5Text, // 34
            Group6Id, // 35
            Group6Text, // 36
            Group7Id, // 37
            Group7Text, // 38
            Group8Id, // 39
            Group8Text, // 40
            Group9Id, // 41
            Group9Text, // 42
            Group10Id, // 43
            Group10Text, // 44
            AssessmentGuid, // 45
            AssessmentSessionLocationId, // 46
            AssessmentSessionLocation, // 47
            AssessmentAdministrationFinishDate,
            AssessmentYear,
            AssessmentType,
            AssessmentAcademicSubject,
            AssessmentLevelForWhichDesigned,
            AssessmentSubtestResultScoreValue,
            AssessmentSubtestMinimumValue,
            AssessmentSubtestMaximumValue,
            AssessmentPerformanceLevelIdentifier,
            AssessmentSubtestResultScoreClaim1Value,
            AssessmentSubtestClaim1MinimumValue,
            AssessmentSubtestClaim1MaximumValue,
            AssessmentClaim1PerformanceLevelIdentifier,
            AssessmentSubtestResultScoreClaim2Value,
            AssessmentSubtestClaim2MinimumValue,
            AssessmentSubtestClaim2MaximumValue,
            AssessmentClaim2PerformanceLevelIdentifier,
            AssessmentSubtestResultScoreClaim3Value,
            AssessmentSubtestClaim3MinimumValue,
            AssessmentSubtestClaim3MaximumValue,
            AssessmentClaim3PerformanceLevelIdentifier,
            AssessmentSubtestResultScoreClaim4Value,
            AssessmentSubtestClaim4MinimumValue,
            AssessmentSubtestClaim4MaximumValue,
            AssessmentClaim4PerformanceLevelIdentifier,
            AccommodationAmericanSignLanguage, // 73
            AccommodationBraille,
            AccommodationClosedCaptioning,
            AccommodationTextToSpeech,
            AccommodationAbacus,
            AccommodationAlternateResponseOptions,
            AccommodationCalculator,
            AccommodationMultiplicationTable,
            AccommodationPrintOnDemand,
            AccommodationPrintOnDemandItems,
            AccommodationReadAloud,
            AccommodationScribe,
            AccommodationSpeechToText,
            AccommodationStreamlineMode,
            AccommodationNoiseBuffer
        };
        static readonly int DwStudentFieldNamesCount = Enum.GetNames(typeof(DwStudentFieldNames)).Length;

        // Student Field names for the "all" format
        enum AllStudentFieldNames : int
        {
            AssessmentId,
            AssessmentName,
            Subject,
            DeliveryMode,
            TestGrade,
            AssessmentType,
            SchoolYear,
            AssessmentVersion,
            StudentIdentifier,
            AlternateSSID,
            FirstName,
            MiddleName,
            LastOrSurname,
            Birthdate,
            GradeLevelWhenAssessed,
            Sex,
            HispanicOrLatinoEthnicity,
            AmericanIndianOrAlaskaNative,
            Asian,
            BlackOrAfricanAmerican,
            White,
            NativeHataiianOrOtherPacificIslander,
            DemographRaceTwoOrMoreRaces,
            IDEAIndicator,
            LEPStatus,
            Section504Status,
            EconomicDisadvantageStatus,
            LanguageCode,
            EnglishLanguageProficiencyLevel,
            MigrantStatus,
            FirstEntryDateIntoUSSchool,
            LimitedEnglishProficiencyEntryDate,
            LEPExitDate,
            TitleIIILanguageInstructionProgramType,
            PrimaryDisabilityType,
            StateAbbreviation,
            DistrictId,
            DistrictName,
            SchoolId,
            SchoolName,
            StudentGroupNames,
            TestOpportunityId,
            TestOpportunityAltId,
            StartDateTime,
            SubmitDateTime,
            Status,
            StatusDateTime,
            AdministrationCondition,
            Completeness,
            AccessibilityCodes,
            NumberOfResponses,
            FieldTestCount,
            PauseCount,
            GracePeriodRestarts,
            AbnormalStarts,
            OpportunityCount,
            TestSessionId,
            UserAgent,
            ScaleScore,
            ScaleScoreStandardError,
            ScaleScoreAchievementLevel,
            Claim1Score,
            Claim1ScoreStandardError,
            Claim1ScoreAchievementLevel,
            Claim2Score,
            Claim2ScoreStandardError,
            Claim2ScoreAchievementLevel,
            Claim3Score,
            Claim3ScoreStandardError,
            Claim3ScoreAchievementLevel,
            Claim4Score,
            Claim4ScoreStandardError,
            Claim4ScoreAchievementLevel,
            OverallTheta,
            OverallThetaStandardError
        };
        static readonly int AllStudentFieldNamesCount = Enum.GetNames(typeof(AllStudentFieldNames)).Length;


        // Item Field names for the Data Warehouse format
        enum DwItemFieldNames : int
        {
            key, // ItemId
            studentId, // May be StudentIdentifier or AlternateSSID
            segmentId,
            position,
            clientId,
            operational,
            isSelected,
            format,
            score,
            scoreStatus,
            adminDate,
            numberVisits,
            strand,
            contentLevel,
            pageNumber,
            pageVisits,
            pageTime,
            dropped,
        };
        static readonly int DwItemFieldNamesCount = Enum.GetNames(typeof(DwItemFieldNames)).Length;

        // Item Field names for the "All Fields" format
        enum AllItemFieldNames : int
        {
            StudentIdentifier, // May be StudentIdentifier or AlternateSSID
            TestOpportunityId,
            SegmentId,
            ItemId,
            ItemPosition,
            Strand,
            ContentLevel,
            FieldTest,
            Dropped,
            ItemType,
            AdminDateTime,
            Submitted,
            SubmitDateTime,
            NumberOfVisits,
            ResponseDuration,
            Score,
            ScoreStatus,
            ScoreDimension,
            ResponseContentType,
            ResponseValue
        };
        static readonly int AllItemFieldNamesCount = Enum.GetNames(typeof(AllItemFieldNames)).Length;

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
        static XPathExpression sXp_LanguageCode = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='LanguageCode' and @context='FINAL']/@value");
        static XPathExpression sXp_EnglishLanguageProficiencyLevel = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[(@name='EnglishLanguageProficiencyLevel' or @name='EnglishLanguageProficiencLevel') and @context='FINAL']/@value");
        static XPathExpression sXp_FirstEntryDateIntoUSSchool = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='FirstEntryDateIntoUSSchool' and @context='FINAL']/@value");
        static XPathExpression sXp_LimitedEnglishProficiencyEntryDate = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='LimitedEnglishProficiencyEntryDate' and @context='FINAL']/@value");
        static XPathExpression sXp_LEPExitDate = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='LEPExitDate' and @context='FINAL']/@value");
        static XPathExpression sXp_TitleIIILanguageInstructionProgramType = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='TitleIIILanguageInstructionProgramType' and @context='FINAL']/@value");
        static XPathExpression sXp_PrimaryDisabilityType = XPathExpression.Compile("/TDSReport/Examinee/ExamineeAttribute[@name='PrimaryDisabilityType' and @context='FINAL']/@value");
        static XPathExpression sXp_StudentGroupName = XPathExpression.Compile("/TDSReport/Examinee/ExamineeRelationship[@name='StudentGroupName' and @context='FINAL']/@value");
        static XPathExpression sXp_AssessmentId = XPathExpression.Compile("/TDSReport/Test/@testId");
        static XPathExpression sXp_AssessmentName = XPathExpression.Compile("/TDSReport/Test/@name");
        static XPathExpression sXp_DeliveryMode = XPathExpression.Compile("/TDSReport/Test/@mode");
        static XPathExpression sXp_AssessmentVersion = XPathExpression.Compile("/TDSReport/Test/@assessmentVersion");
        static XPathExpression sXp_TestSessionId = XPathExpression.Compile("/TDSReport/Opportunity/@sessionId");
        static XPathExpression sXp_SchoolYear = XPathExpression.Compile("/TDSReport/Test/@academicYear");
        static XPathExpression sXp_AssessmentType = XPathExpression.Compile("/TDSReport/Test/@assessmentType");
        static XPathExpression sXp_Subject = XPathExpression.Compile("/TDSReport/Test/@subject");
        static XPathExpression sXp_TestGrade = XPathExpression.Compile("/TDSReport/Test/@grade");
        static XPathExpression sXp_ScaleScore = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@value");
        static XPathExpression sXp_ScaleScoreStandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@standardError");
        static XPathExpression sXp_ScaleScoreAchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and (@measureLabel='PerformanceLevel' or @measurelabel='PerformanceLevel')]/@value");
        // Claim 1 for ELA is labeled "SOCK_R" for math it's labeled "1"
        static XPathExpression sXp_ClaimScore1 = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_R' or @measureOf='1' or @measureOf='Claim1') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@value");
        static XPathExpression sXp_ClaimScore1StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_R' or @measureOf='1' or @measureOf='Claim1') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@standardError");
        static XPathExpression sXp_ClaimScore1AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_R' or @measureOf='1' or @measureOf='Claim1') and (@measureLabel='PerformanceLevel' or @measurelabel='PerformanceLevel')]/@value");
        // Claim 2 for ELA is labeled "2-W" for math it's labeled "SOCK_2"
        static XPathExpression sXp_ClaimScore2 = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='2-W' or @measureOf='SOCK_2' or @measureOf='Claim2') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@value");
        static XPathExpression sXp_ClaimScore2StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='2-W' or @measureOf='SOCK_2' or @measureOf='Claim2') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@standardError");
        static XPathExpression sXp_ClaimScore2AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='2-W' or @measureOf='SOCK_2' or @measureOf='Claim2') and (@measureLabel='PerformanceLevel' or @measurelabel='PerformanceLevel')]/@value");
        // Claim 3 for ELA is labeled "SOCK_LS" for math it's labeled "3"
        static XPathExpression sXp_ClaimScore3 = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_LS' or @measureOf='3' or @measureOf='Claim3') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@value");
        static XPathExpression sXp_ClaimScore3StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_LS' or @measureOf='3' or @measureOf='Claim3') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@standardError");
        static XPathExpression sXp_ClaimScore3AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='SOCK_LS' or @measureOf='3' or @measureOf='Claim3') and (@measureLabel='PerformanceLevel' or @measurelabel='PerformanceLevel')]/@value");
        // Claim 4 for ELA is labeled "4-CR" for math it doesn't exist
        static XPathExpression sXp_ClaimScore4 = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='4-CR' or @measureOf='Claim4') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@value");
        static XPathExpression sXp_ClaimScore4StandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='4-CR' or @measureOf='Claim4') and (@measureLabel='ScaleScore' or @measurelabel='ScaleScore')]/@standardError");
        static XPathExpression sXp_ClaimScore4AchievementLevel = XPathExpression.Compile("/TDSReport/Opportunity/Score[(@measureOf='4-CR' or @measureOf='Claim4') and (@measureLabel='PerformanceLevel' or @measurelabel='PerformanceLevel')]/@value");
        static XPathExpression sXp_OverallTheta = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and (@measurelabel='Theta' or @measurelabel='Pre-LOSS/HOSS theta')]/@value");
        static XPathExpression sXp_OverallThetaStandardError = XPathExpression.Compile("/TDSReport/Opportunity/Score[@measureOf='Overall' and (@measurelabel='Theta' or @measurelabel='Pre-LOSS/HOSS theta')]/@standardError");
        // Matches all accessibility codes
        static XPathExpression sXP_AccessibilityCodes = XPathExpression.Compile("/TDSReport/Opportunity/Accommodation/@code");

        static XPathExpression sXp_TestOpportunityId = XPathExpression.Compile("/TDSReport/Opportunity/@key");
        static XPathExpression sXp_TestOpportunityAltId = XPathExpression.Compile("/TDSReport/Opportunity/@oppId");
        static XPathExpression sXp_StartDateTime = XPathExpression.Compile("/TDSReport/Opportunity/@startDate");
        static XPathExpression sXp_SubmitDateTime = XPathExpression.Compile("/TDSReport/Opportunity/@dateCompleted");
        static XPathExpression sXp_Status = XPathExpression.Compile("/TDSReport/Opportunity/@status");
        static XPathExpression sXp_StatusDateTime = XPathExpression.Compile("/TDSReport/Opportunity/@statusDate");
        static XPathExpression sXp_AdministrationCondition_0 = XPathExpression.Compile("/TDSReport/Opportunity/@administrationCondition");
        static XPathExpression sXp_AdministrationCondition_1 = XPathExpression.Compile("/TDSReport/Opportunity/@validity");
        static XPathExpression sXp_Completeness_0 = XPathExpression.Compile("/TDSReport/Opportunity/@completeness");
        static XPathExpression sXp_Completeness_1 = XPathExpression.Compile("/TDSReport/Opportunity/@completeStatus");
        static XPathExpression sXp_NumberOfResponses = XPathExpression.Compile("/TDSReport/Opportunity/@itemCount");
        static XPathExpression sXp_FieldTestCount = XPathExpression.Compile("/TDSReport/Opportunity/@ftCount");
        static XPathExpression sXp_PauseCount = XPathExpression.Compile("/TDSReport/Opportunity/@pauseCount");
        static XPathExpression sXp_GracePeriodRestarts = XPathExpression.Compile("/TDSReport/Opportunity/@gracePeriodRestarts");
        static XPathExpression sXp_AbnormalStarts = XPathExpression.Compile("/TDSReport/Opportunity/@abnormalStarts");
        static XPathExpression sXp_OpportunityCount = XPathExpression.Compile("/TDSReport/Opportunity/@opportunity");
        static XPathExpression sXp_UserAgent = XPathExpression.Compile("/TDSReport/Opportunity/@assessmentParticipantSessionPlatformUserAgent");

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
        static XPathExpression sXp_ItemSubmitDateTime = XPathExpression.Compile("Response/@date");
        static XPathExpression sXp_ScoreDimension = XPathExpression.Compile("ScoreInfo/@scoreDimension");
        static XPathExpression sXp_ResponseContentType = XPathExpression.Compile("@mimeType");
        static XPathExpression sXp_Response = XPathExpression.Compile("Response");

        static XPathExpression sXp_SubScoreInfo = XPathExpression.Compile("ScoreInfo/SubScoreList/ScoreInfo");

        static XPathExpression sXpSubScoreScore = XPathExpression.Compile("@scorePoint");
        static XPathExpression sXp_SubScoreStatus = XPathExpression.Compile("@scoreStatus");
        static XPathExpression sXp_SubScoreDimension = XPathExpression.Compile("@scoreDimension");

        static Dictionary<string, int> sAccessibilityCodeMapping;

        static ToCsvProcessor()
        {
            // Only include codes that make the feature available.
            // So, leave out TDS_ASL0, TDS_ClosedCap0, TDS_TTS0, 
            sAccessibilityCodeMapping = new Dictionary<string, int>();
            sAccessibilityCodeMapping.Add("TDS_ASL1", (int)DwStudentFieldNames.AccommodationAmericanSignLanguage); // American Sign Language
            sAccessibilityCodeMapping.Add("TDS_ClosedCap1", (int)DwStudentFieldNames.AccommodationClosedCaptioning); // Closed Captioning
            sAccessibilityCodeMapping.Add("ENU-Braille", (int)DwStudentFieldNames.AccommodationBraille); // Braille
            sAccessibilityCodeMapping.Add("TDS_PoD_Stim", (int)DwStudentFieldNames.AccommodationPrintOnDemand); // Print on Demand Stimuli
            sAccessibilityCodeMapping.Add("TDS_PoD_Item", (int)DwStudentFieldNames.AccommodationPrintOnDemandItems); // Print on Demand Item
            sAccessibilityCodeMapping.Add("TDS_TS_Accessibility", (int)DwStudentFieldNames.AccommodationStreamlineMode); // Streamline
            sAccessibilityCodeMapping.Add("TDS_TTS_Item", (int)DwStudentFieldNames.AccommodationTextToSpeech); // Text to Speech
            sAccessibilityCodeMapping.Add("TDS_TTS_Stim", (int)DwStudentFieldNames.AccommodationTextToSpeech);
            sAccessibilityCodeMapping.Add("TDS_TTS_Stim&TDS_TTS_Item", (int)DwStudentFieldNames.AccommodationTextToSpeech);
            sAccessibilityCodeMapping.Add("NEA_AR", (int)DwStudentFieldNames.AccommodationAlternateResponseOptions); // Non-Embedded Alternate Response Options
            sAccessibilityCodeMapping.Add("NEA_RA_Stimuli", (int)DwStudentFieldNames.AccommodationReadAloud); // Non-Embedded Read Aloud
            sAccessibilityCodeMapping.Add("NEA_SC_WritItems", (int)DwStudentFieldNames.AccommodationScribe); // Non-Embedded Scribe
            sAccessibilityCodeMapping.Add("NEA_STT", (int)DwStudentFieldNames.AccommodationSpeechToText); // Non-Embedded Speech to Text
            sAccessibilityCodeMapping.Add("NEA_Abacus", (int)DwStudentFieldNames.AccommodationAbacus); // Non-Embedded Abacus
            sAccessibilityCodeMapping.Add("NEA_Calc", (int)DwStudentFieldNames.AccommodationCalculator); // Non-Embedded Calculator
            sAccessibilityCodeMapping.Add("NEA_MT", (int)DwStudentFieldNames.AccommodationMultiplicationTable); // Non-Embedded Multiplication Table
            sAccessibilityCodeMapping.Add("NEA_NoiseBuf", (int)DwStudentFieldNames.AccommodationNoiseBuffer); // Non-Embedded Noise Buffer
        }
      
        Parse.CsvWriter m_sWriter;
        Parse.CsvWriter m_iWriter;
        byte[] m_hashKey;

        public ToCsvProcessor(string osFilename, string oiFilename, OutputFormat outputFormat)
        {
            OutputFormat = outputFormat;
#if !DEBUG
            if (File.Exists(osFilename)) throw new ApplicationException(string.Format("Output file, '{0}' already exists.", osFilename));
            if (File.Exists(oiFilename)) throw new ApplicationException(string.Format("Output file, '{0}' already exists.", oiFilename));
#endif
            if (osFilename != null)
            {
                m_sWriter = new Parse.CsvWriter(osFilename, false);
                string[] fieldNames = (OutputFormat == OutputFormat.Dw) ? Enum.GetNames(typeof(DwStudentFieldNames)) : Enum.GetNames(typeof(AllStudentFieldNames));
                m_sWriter.Write(fieldNames);
            }
            if (oiFilename != null)
            {
                m_iWriter = new Parse.CsvWriter(oiFilename, false);
                string[] fieldNames = (OutputFormat == OutputFormat.Dw) ? Enum.GetNames(typeof(DwItemFieldNames)) : Enum.GetNames(typeof(AllItemFieldNames));
                m_iWriter.Write(fieldNames);
            }
        }

        public DIDFlags DIDFlags { get; set; }

        public OutputFormat OutputFormat { get; private set; }

        public int MaxResponse { get; set; }

        public bool NotExcel { get; set; }

        static readonly UTF8Encoding UTF8NoByteOrderMark = new UTF8Encoding(false);

        public string HashPassPhrase
        {
            set
            {
                if (value == null)
                {
                    m_hashKey = null;
                }
                else
                {
                    SHA1 sha = new SHA1CryptoServiceProvider();
                    byte[] pfb = UTF8NoByteOrderMark.GetBytes(value);
                    m_hashKey = sha.ComputeHash(pfb);
                }
            }
        }

        public void ProcessResult(Stream input)
        {
            XPathDocument doc = new XPathDocument(input);
            XPathNavigator nav = doc.CreateNavigator();

            // Retrieve common fields and those that may be manipulated by de-identification
            string studentId = nav.Eval(sXp_StudentIdentifier);
            string alternateSSID = nav.Eval(sXp_AlternateSSID);
            if (string.IsNullOrEmpty(studentId)) studentId = alternateSSID;
            string testOpportunityId = nav.Eval(sXp_TestOpportunityId);

            string firstName = nav.Eval(sXp_FirstName);
            string middleName = nav.Eval(sXp_MiddleName);
            string lastOrSurname = nav.Eval(sXp_LastOrSurname);

            string birthdate = NormalizeDate(nav.Eval(sXp_Birthdate));

            string sex = NormalizeSex(nav.Eval(sXp_Sex));
            string hispanicOrLatinoEthnicity = nav.Eval(sXp_HispanicOrLatinoEthnicity);
            string americanIndianOrAlaskaNative = nav.Eval(sXp_AmericanIndianOrAlaskaNative);
            string asian = nav.Eval(sXp_Asian);
            string blackOrAfricanAmerican = nav.Eval(sXp_BlackOrAfricanAmerican);
            string nativeHawaiianOrOtherPacificIslander = nav.Eval(sXp_NativeHawaiianOrOtherPacificIslander);
            string white = nav.Eval(sXp_White);
            string demographicRaceTwoOrMoreRaces = nav.Eval(sXp_DemographicRaceTwoOrMoreRaces);
            string IDEAIndicator = nav.Eval(sXp_IDEAIndicator);
            string LEPStatus = nav.Eval(sXp_LEPStatus);
            string section504Status = nav.Eval(sXp_Section504Status);
            string economicDisadvantageStatus = nav.Eval(sXp_EconomicDisadvantageStatus);
            string languageCode = nav.Eval(sXp_LanguageCode);
            string englishLanguageProficiencyLevel = nav.Eval(sXp_EnglishLanguageProficiencyLevel);
            string firstEntryDateIntoUSSchool = NormalizeDate(nav.Eval(sXp_FirstEntryDateIntoUSSchool));
            string limitedEnglishProficiencyEntryDate = NormalizeDate(nav.Eval(sXp_LimitedEnglishProficiencyEntryDate));
            string LEPExitDate = NormalizeDate(nav.Eval(sXp_LEPExitDate));
            string titleIIILanguageInstructionProgramType = nav.Eval(sXp_TitleIIILanguageInstructionProgramType);
            string primaryDisabilityType = nav.Eval(sXp_PrimaryDisabilityType);
            string migrantStatus = nav.Eval(sXp_MigrantStatus);

            string districtId = nav.Eval(sXp_DistrictId);
            string districtName = nav.Eval(sXp_DistrictName);
            string schoolId = nav.Eval(sXp_SchoolId);
            string schoolName = nav.Eval(sXp_SchoolName);
            string testSessionId = nav.Eval(sXp_TestSessionId);

            // Get student groups (should be 10 or fewer)
            string[] groups = null;
            {
                List<string> groupList = new List<string>();
                XPathNodeIterator nodes = nav.Select(sXp_StudentGroupName);
                while (nodes.MoveNext())
                {
                    groupList.Add(nodes.Current.ToString().Replace(';', '-'));
                    if (nodes.Count >= 10) break;
                }
                groups = groupList.ToArray();
            }

            // If hash key provided, cryptographically hash the student ID
            if (m_hashKey != null)
            {
                HMACSHA1 hmac = new HMACSHA1(m_hashKey);
                byte[] bid = UTF8NoByteOrderMark.GetBytes(studentId);
                byte[] hash = hmac.ComputeHash(bid);
                alternateSSID = ByteArrayToHexString(hash);
            }

            // De-identify if specified
            if ((this.DIDFlags & DIDFlags.Id) != 0)
            {
                // Student ID is a special case. De-identification means to substitute
                // the AlternateSSID for the student Id.
                studentId = alternateSSID;
            }
            if ((this.DIDFlags & DIDFlags.Name) != 0)
            {
                firstName = string.Empty;
                middleName = string.Empty;
                lastOrSurname = string.Empty;
            }
            if ((this.DIDFlags & DIDFlags.Birthdate) != 0)
            {
                birthdate = string.Empty;
            }
            if ((this.DIDFlags & DIDFlags.Demographics) != 0)
            {
                sex = string.Empty;
                hispanicOrLatinoEthnicity = string.Empty;
                americanIndianOrAlaskaNative = string.Empty;
                asian = string.Empty;
                blackOrAfricanAmerican = string.Empty;
                nativeHawaiianOrOtherPacificIslander = string.Empty;
                white = string.Empty;
                demographicRaceTwoOrMoreRaces = string.Empty;
                IDEAIndicator = string.Empty;
                LEPStatus = string.Empty;
                section504Status = string.Empty;
                economicDisadvantageStatus = string.Empty;
                languageCode = string.Empty;
                englishLanguageProficiencyLevel = string.Empty;
                firstEntryDateIntoUSSchool = string.Empty;
                limitedEnglishProficiencyEntryDate = string.Empty;
                LEPExitDate = string.Empty;
                titleIIILanguageInstructionProgramType = string.Empty;
                primaryDisabilityType = string.Empty;

                migrantStatus = string.Empty;
            }
            if ((this.DIDFlags & DIDFlags.School) != 0)
            {
                districtId = string.Empty;
                districtName = string.Empty;
                schoolId = string.Empty;
                schoolName = string.Empty;
                testSessionId = string.Empty;
            }
            if (this.DIDFlags != DIDFlags.None)
            {
                groups = new string[0];
            }

            // Retrieve the student fields
            string[] studentFields;

            if (OutputFormat == OutputFormat.Dw)
            {
                studentFields = new string[DwStudentFieldNamesCount];

                studentFields[(int)DwStudentFieldNames.StateAbbreviation] = nav.Eval(sXp_StateAbbreviation);
                studentFields[(int)DwStudentFieldNames.ResponsibleDistrictIdentifier] = districtId;
                studentFields[(int)DwStudentFieldNames.OrganizationName] = districtName;
                studentFields[(int)DwStudentFieldNames.ResponsibleSchoolIdentifier] = schoolId;
                studentFields[(int)DwStudentFieldNames.NameOfInstitution] = schoolName;
                studentFields[(int)DwStudentFieldNames.StudentIdentifier] = studentId;
                studentFields[(int)DwStudentFieldNames.ExternalSSID] = alternateSSID;
                studentFields[(int)DwStudentFieldNames.FirstName] = firstName;
                studentFields[(int)DwStudentFieldNames.MiddleName] = middleName;
                studentFields[(int)DwStudentFieldNames.LastOrSurname] = lastOrSurname;
                studentFields[(int)DwStudentFieldNames.Sex] = sex;
                studentFields[(int)DwStudentFieldNames.Birthdate] = birthdate;
                studentFields[(int)DwStudentFieldNames.GradeLevelWhenAssessed] = NormalizeGrade(nav.Eval(sXp_GradeLevelWhenAssessed));
                studentFields[(int)DwStudentFieldNames.HispanicOrLatinoEthnicity] = hispanicOrLatinoEthnicity;
                studentFields[(int)DwStudentFieldNames.AmericanIndianOrAlaskaNative] = americanIndianOrAlaskaNative;
                studentFields[(int)DwStudentFieldNames.Asian] = asian;
                studentFields[(int)DwStudentFieldNames.BlackOrAfricanAmerican] = blackOrAfricanAmerican;
                studentFields[(int)DwStudentFieldNames.NativeHawaiianOrOtherPacificIslander] = nativeHawaiianOrOtherPacificIslander;
                studentFields[(int)DwStudentFieldNames.White] = white;
                studentFields[(int)DwStudentFieldNames.DemographicRaceTwoOrMoreRaces] = demographicRaceTwoOrMoreRaces;
                studentFields[(int)DwStudentFieldNames.IDEAIndicator] = IDEAIndicator;
                studentFields[(int)DwStudentFieldNames.LEPStatus] = LEPStatus;
                studentFields[(int)DwStudentFieldNames.Section504Status] = section504Status;
                studentFields[(int)DwStudentFieldNames.EconomicDisadvantageStatus] = economicDisadvantageStatus;
                studentFields[(int)DwStudentFieldNames.MigrantStatus] = migrantStatus;

                for (int i = 0; i < 10; ++i)
                {
                    if (i < groups.Length)
                    {
                        studentFields[i * 2 + (int)DwStudentFieldNames.Group1Id] = groups[i];
                        studentFields[i * 2 + (int)DwStudentFieldNames.Group1Text] = groups[i];
                    }
                    else
                    {
                        studentFields[i * 2 + (int)DwStudentFieldNames.Group1Id] = string.Empty;
                        studentFields[i * 2 + (int)DwStudentFieldNames.Group1Text] = string.Empty;
                    }
                }

                studentFields[(int)DwStudentFieldNames.AssessmentGuid] = nav.Eval(sXp_AssessmentId);
                studentFields[(int)DwStudentFieldNames.AssessmentSessionLocationId] = testSessionId;
                studentFields[(int)DwStudentFieldNames.AssessmentSessionLocation] = string.Empty; // AssessmentLocation
                studentFields[(int)DwStudentFieldNames.AssessmentAdministrationFinishDate] = string.Empty; // AssessmentAdministrationFinishDate
                studentFields[(int)DwStudentFieldNames.AssessmentYear] = nav.Eval(sXp_SchoolYear);
                studentFields[(int)DwStudentFieldNames.AssessmentType] = nav.Eval(sXp_AssessmentType);
                studentFields[(int)DwStudentFieldNames.AssessmentAcademicSubject] = nav.Eval(sXp_Subject);
                studentFields[(int)DwStudentFieldNames.AssessmentLevelForWhichDesigned] = NormalizeGrade(nav.Eval(sXp_TestGrade));
                ProcessScoresDw(nav, sXp_ScaleScore, sXp_ScaleScoreStandardError, sXp_ScaleScoreAchievementLevel, studentFields, (int)DwStudentFieldNames.AssessmentSubtestResultScoreValue);
                ProcessScoresDw(nav, sXp_ClaimScore1, sXp_ClaimScore1StandardError, sXp_ClaimScore1AchievementLevel, studentFields, (int)DwStudentFieldNames.AssessmentSubtestResultScoreClaim1Value);
                ProcessScoresDw(nav, sXp_ClaimScore2, sXp_ClaimScore2StandardError, sXp_ClaimScore2AchievementLevel, studentFields, (int)DwStudentFieldNames.AssessmentSubtestResultScoreClaim2Value);
                ProcessScoresDw(nav, sXp_ClaimScore3, sXp_ClaimScore3StandardError, sXp_ClaimScore3AchievementLevel, studentFields, (int)DwStudentFieldNames.AssessmentSubtestResultScoreClaim3Value);
                ProcessScoresDw(nav, sXp_ClaimScore4, sXp_ClaimScore4StandardError, sXp_ClaimScore4AchievementLevel, studentFields, (int)DwStudentFieldNames.AssessmentSubtestResultScoreClaim4Value);

                // Preload accommodation fields with "3" (accessibility feature not made available)
                for (int i = (int)DwStudentFieldNames.AccommodationAmericanSignLanguage; i <= (int)DwStudentFieldNames.AccommodationNoiseBuffer; ++i) studentFields[i] = "3";

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
                                studentFields[fieldIndex] = "6"; // Set code to 6 accessibility feature made available.
                            }
                        }
                    }
                }
            }

            // Load up the "all" student fields format
            else
            {
                studentFields = new string[AllStudentFieldNamesCount];
                
                studentFields[(int)AllStudentFieldNames.AssessmentId] = nav.Eval(sXp_AssessmentId);
                studentFields[(int)AllStudentFieldNames.AssessmentName] = nav.Eval(sXp_AssessmentName);
                studentFields[(int)AllStudentFieldNames.Subject] = nav.Eval(sXp_Subject);
                studentFields[(int)AllStudentFieldNames.DeliveryMode] = nav.Eval(sXp_DeliveryMode);
                studentFields[(int)AllStudentFieldNames.TestGrade] = nav.Eval(sXp_TestGrade);
                studentFields[(int)AllStudentFieldNames.AssessmentType] = nav.Eval(sXp_AssessmentType);
                studentFields[(int)AllStudentFieldNames.SchoolYear] = nav.Eval(sXp_SchoolYear);
                studentFields[(int)AllStudentFieldNames.AssessmentVersion] = nav.Eval(sXp_AssessmentVersion);
                studentFields[(int)AllStudentFieldNames.StudentIdentifier] = NotExcel ? studentId : studentId + "\t"; // Tab causes Excel to treat this as text, not number
                studentFields[(int)AllStudentFieldNames.AlternateSSID] = NotExcel ? alternateSSID : alternateSSID + "\t";
                studentFields[(int)AllStudentFieldNames.FirstName] = firstName;
                studentFields[(int)AllStudentFieldNames.MiddleName] = middleName;
                studentFields[(int)AllStudentFieldNames.LastOrSurname] = lastOrSurname;
                studentFields[(int)AllStudentFieldNames.Birthdate] = birthdate;
                studentFields[(int)AllStudentFieldNames.GradeLevelWhenAssessed] = NormalizeGrade(nav.Eval(sXp_GradeLevelWhenAssessed));
                studentFields[(int)AllStudentFieldNames.Sex] = sex;
                studentFields[(int)AllStudentFieldNames.HispanicOrLatinoEthnicity] = hispanicOrLatinoEthnicity;
                studentFields[(int)AllStudentFieldNames.AmericanIndianOrAlaskaNative] = americanIndianOrAlaskaNative;
                studentFields[(int)AllStudentFieldNames.Asian] = asian;
                studentFields[(int)AllStudentFieldNames.BlackOrAfricanAmerican] = blackOrAfricanAmerican;
                studentFields[(int)AllStudentFieldNames.White] = white;
                studentFields[(int)AllStudentFieldNames.NativeHataiianOrOtherPacificIslander] = nativeHawaiianOrOtherPacificIslander;
                studentFields[(int)AllStudentFieldNames.DemographRaceTwoOrMoreRaces] = demographicRaceTwoOrMoreRaces;
                studentFields[(int)AllStudentFieldNames.IDEAIndicator] = IDEAIndicator;
                studentFields[(int)AllStudentFieldNames.LEPStatus] = LEPStatus;
                studentFields[(int)AllStudentFieldNames.Section504Status] = section504Status;
                studentFields[(int)AllStudentFieldNames.EconomicDisadvantageStatus] = economicDisadvantageStatus;
                studentFields[(int)AllStudentFieldNames.LanguageCode] = languageCode;
                studentFields[(int)AllStudentFieldNames.EnglishLanguageProficiencyLevel] = englishLanguageProficiencyLevel;
                studentFields[(int)AllStudentFieldNames.MigrantStatus] = migrantStatus;
                studentFields[(int)AllStudentFieldNames.FirstEntryDateIntoUSSchool] = firstEntryDateIntoUSSchool;
                studentFields[(int)AllStudentFieldNames.LimitedEnglishProficiencyEntryDate] = limitedEnglishProficiencyEntryDate;
                studentFields[(int)AllStudentFieldNames.LEPExitDate] = LEPExitDate;
                studentFields[(int)AllStudentFieldNames.TitleIIILanguageInstructionProgramType] = titleIIILanguageInstructionProgramType;
                studentFields[(int)AllStudentFieldNames.PrimaryDisabilityType] = primaryDisabilityType;
                studentFields[(int)AllStudentFieldNames.StateAbbreviation] = nav.Eval(sXp_StateAbbreviation);
                studentFields[(int)AllStudentFieldNames.DistrictId] = districtId;
                studentFields[(int)AllStudentFieldNames.DistrictName] = districtName;
                studentFields[(int)AllStudentFieldNames.SchoolId] = schoolId;
                studentFields[(int)AllStudentFieldNames.SchoolName] = schoolName;
                studentFields[(int)AllStudentFieldNames.StudentGroupNames] = string.Join(";", groups);
                studentFields[(int)AllStudentFieldNames.TestOpportunityId] = testOpportunityId;
                studentFields[(int)AllStudentFieldNames.TestOpportunityAltId] = nav.Eval(sXp_TestOpportunityAltId);
                studentFields[(int)AllStudentFieldNames.StartDateTime] = NormalizeDateTime(nav.Eval(sXp_StartDateTime));
                studentFields[(int)AllStudentFieldNames.SubmitDateTime] = NormalizeDateTime(nav.Eval(sXp_SubmitDateTime));
                studentFields[(int)AllStudentFieldNames.Status] = nav.Eval(sXp_Status);
                studentFields[(int)AllStudentFieldNames.StatusDateTime] = nav.Eval(sXp_StatusDateTime);
                studentFields[(int)AllStudentFieldNames.AdministrationCondition] = nav.EvalFirstMatch(sXp_AdministrationCondition_0, sXp_AdministrationCondition_1);
                studentFields[(int)AllStudentFieldNames.Completeness] = nav.EvalFirstMatch(sXp_Completeness_0, sXp_Completeness_1);

                studentFields[(int)AllStudentFieldNames.AccessibilityCodes] = CompileAccessibilityCodes(nav);

                studentFields[(int)AllStudentFieldNames.NumberOfResponses] = nav.Eval(sXp_NumberOfResponses);
                studentFields[(int)AllStudentFieldNames.FieldTestCount] = nav.Eval(sXp_FieldTestCount);
                studentFields[(int)AllStudentFieldNames.PauseCount] = nav.Eval(sXp_PauseCount);
                studentFields[(int)AllStudentFieldNames.GracePeriodRestarts] = nav.Eval(sXp_GracePeriodRestarts);
                studentFields[(int)AllStudentFieldNames.AbnormalStarts] = nav.Eval(sXp_AbnormalStarts);
                studentFields[(int)AllStudentFieldNames.OpportunityCount] = nav.Eval(sXp_OpportunityCount);
                studentFields[(int)AllStudentFieldNames.TestSessionId] = nav.Eval(sXp_TestSessionId);
                studentFields[(int)AllStudentFieldNames.UserAgent] = nav.Eval(sXp_UserAgent);
                ProcessScoresAll(nav, sXp_ScaleScore, sXp_ScaleScoreStandardError, sXp_ScaleScoreAchievementLevel, studentFields, (int)AllStudentFieldNames.ScaleScore);
                ProcessScoresAll(nav, sXp_ClaimScore1, sXp_ClaimScore1StandardError, sXp_ClaimScore1AchievementLevel, studentFields, (int)AllStudentFieldNames.Claim1Score);
                ProcessScoresAll(nav, sXp_ClaimScore2, sXp_ClaimScore2StandardError, sXp_ClaimScore2AchievementLevel, studentFields, (int)AllStudentFieldNames.Claim2Score);
                ProcessScoresAll(nav, sXp_ClaimScore3, sXp_ClaimScore3StandardError, sXp_ClaimScore3AchievementLevel, studentFields, (int)AllStudentFieldNames.Claim3Score);
                ProcessScoresAll(nav, sXp_ClaimScore4, sXp_ClaimScore4StandardError, sXp_ClaimScore4AchievementLevel, studentFields, (int)AllStudentFieldNames.Claim4Score);
                studentFields[(int)AllStudentFieldNames.OverallTheta] = nav.Eval(sXp_OverallTheta);
                studentFields[(int)AllStudentFieldNames.OverallThetaStandardError] = nav.Eval(sXp_OverallThetaStandardError);
            }

            // Write one line to the CSV
            if (m_sWriter != null)
                m_sWriter.Write(studentFields);

            // Report item data
            if (m_iWriter != null)
            {
                XPathNodeIterator nodes = nav.Select(sXP_Item);
                while (nodes.MoveNext())
                {
                    // Collect the item fields
                    string[] itemFields;
                    XPathNavigator node = nodes.Current;

                    if (OutputFormat == OutputFormat.Dw)
                    {
                        itemFields = new string[DwItemFieldNamesCount];

                        itemFields[(int)DwItemFieldNames.key] = string.Concat(node.Eval(sXp_ItemBankKey), "-", node.Eval(sXp_ItemKey));
                        itemFields[(int)DwItemFieldNames.studentId] = studentId;
                        itemFields[(int)DwItemFieldNames.segmentId] = node.Eval(sXp_SegmentId);
                        itemFields[(int)DwItemFieldNames.position] = node.Eval(sXp_ItemPosition);
                        itemFields[(int)DwItemFieldNames.clientId] = node.Eval(sXp_ClientId);
                        itemFields[(int)DwItemFieldNames.operational] = node.Eval(sXp_Operational);
                        itemFields[(int)DwItemFieldNames.isSelected] = node.Eval(sXp_IsSelected);
                        itemFields[(int)DwItemFieldNames.format] = node.Eval(sXp_ItemType);
                        itemFields[(int)DwItemFieldNames.score] = node.Eval(sXp_ItemScore);
                        string scoreStatus = node.Eval(sXp_ScoreStatus);
                        itemFields[(int)DwItemFieldNames.scoreStatus] = string.IsNullOrEmpty(scoreStatus) ? "NOTSCORED" : scoreStatus;
                        itemFields[(int)DwItemFieldNames.adminDate] = node.Eval(sXp_AdminDate);
                        itemFields[(int)DwItemFieldNames.numberVisits] = node.Eval(sXp_NumberVisits);
                        itemFields[(int)DwItemFieldNames.strand] = node.Eval(sXp_Strand);
                        itemFields[(int)DwItemFieldNames.contentLevel] = node.Eval(sXp_ContentLevel);
                        itemFields[(int)DwItemFieldNames.pageNumber] = node.Eval(sXp_PageNumber);
                        itemFields[(int)DwItemFieldNames.pageVisits] = node.Eval(sXp_PageVisits);
                        itemFields[(int)DwItemFieldNames.pageTime] = node.Eval(sXp_PageTime);
                        itemFields[(int)DwItemFieldNames.dropped] = node.Eval(sXp_Dropped);

                        // Write one line to the CSV
                        m_iWriter.Write(itemFields);
                    }

                    // All Fields format
                    else
                    {
                        itemFields = new string[AllItemFieldNamesCount];

                        itemFields[(int)AllItemFieldNames.StudentIdentifier] = studentId; // May be StudentIdentifier or AlternateSSID
                        itemFields[(int)AllItemFieldNames.TestOpportunityId] = testOpportunityId;
                        itemFields[(int)AllItemFieldNames.SegmentId] = node.Eval(sXp_SegmentId);
                        itemFields[(int)AllItemFieldNames.ItemId] = string.Concat(node.Eval(sXp_ItemBankKey), "-", node.Eval(sXp_ItemKey));
                        itemFields[(int)AllItemFieldNames.ItemPosition] = node.Eval(sXp_ItemPosition);
                        itemFields[(int)AllItemFieldNames.Strand] = node.Eval(sXp_Strand);
                        itemFields[(int)AllItemFieldNames.ContentLevel] = node.Eval(sXp_ContentLevel);
                        itemFields[(int)AllItemFieldNames.FieldTest] = NormalizeBool(node.Eval(sXp_Operational)) ? "No" : "Yes";
                        itemFields[(int)AllItemFieldNames.Dropped] = NormalizeBool(node.Eval(sXp_Dropped)) ? "Yes" : "No";
                        itemFields[(int)AllItemFieldNames.ItemType] = node.Eval(sXp_ItemType);
                        itemFields[(int)AllItemFieldNames.AdminDateTime] = NormalizeDateTime(node.Eval(sXp_AdminDate));
                        itemFields[(int)AllItemFieldNames.Submitted] = NormalizeBool(node.Eval(sXp_IsSelected)) ? "Yes" : "No";
                        itemFields[(int)AllItemFieldNames.SubmitDateTime] = NormalizeDateTime(node.Eval(sXp_ItemSubmitDateTime));
                        itemFields[(int)AllItemFieldNames.NumberOfVisits] = node.Eval(sXp_NumberVisits);
                        itemFields[(int)AllItemFieldNames.ResponseDuration] = CalcResponseDuration(nav, node);
                        itemFields[(int)AllItemFieldNames.Score] = node.Eval(sXp_ItemScore);
                        itemFields[(int)AllItemFieldNames.ScoreStatus] = node.Eval(sXp_ScoreStatus);
                        itemFields[(int)AllItemFieldNames.ScoreDimension] = node.Eval(sXp_ScoreDimension);
                        itemFields[(int)AllItemFieldNames.ResponseContentType] = node.Eval(sXp_ResponseContentType);

                        string response;
                        XPathNavigator responseNode = node.SelectSingleNode(sXp_Response);
                        response = (responseNode != null) ? responseNode.Value.Trim() : string.Empty;
                        if (MaxResponse != 0 && response.Length > MaxResponse) response = string.Empty;
                        itemFields[(int)AllItemFieldNames.ResponseValue] = response;

                        // Write one line to the CSV
                        m_iWriter.Write(itemFields);

                        // Don't repeat the (possibly bulky) response for each dimension
                        itemFields[(int)AllItemFieldNames.ResponseContentType] = string.Empty;
                        itemFields[(int)AllItemFieldNames.ResponseValue] = string.Empty;

                        // Process any subscores (for multiple scoring dimensions)
                        XPathNodeIterator subScoreInfoIterator = node.Select(sXp_SubScoreInfo);
                        while (subScoreInfoIterator.MoveNext())
                        {
                            XPathNavigator sub = subScoreInfoIterator.Current;
                            string dimension = sub.Eval(sXp_SubScoreDimension);
                            if (string.Equals(dimension, "Initial", StringComparison.OrdinalIgnoreCase) || string.Equals(dimension, "Final", StringComparison.OrdinalIgnoreCase)) continue;

                            itemFields[(int)AllItemFieldNames.Score] = sub.Eval(sXpSubScoreScore);
                            itemFields[(int)AllItemFieldNames.ScoreStatus] = sub.Eval(sXp_SubScoreStatus);
                            itemFields[(int)AllItemFieldNames.ScoreDimension] = dimension;

                            // Write one line to the CSV
                            m_iWriter.Write(itemFields);
                        }
                    }
                }
            }
        }

        static string NormalizeSex(string sex)
        {
            if (sex.Length < 1) return string.Empty;
            if (sex[0] == 'M' || sex[0] == 'm') return "Male";
            if (sex[0] == 'F' || sex[0] == 'f') return "Female";
            return string.Empty;
        }

        static string NormalizeGrade(string grade)
        {
            return grade;
            /*
            int g;
            if (int.TryParse(grade, out g))
            {
                return g.ToString("D2");
            }
            return grade;
            */
        }

        static string NormalizeDate(string date)
        {
            DateTime d;
            if (DateTime.TryParseExact(date, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal|DateTimeStyles.NoCurrentDateDefault|DateTimeStyles.AllowWhiteSpaces, out d))
            {
                return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            if (DateTime.TryParse(date, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal|DateTimeStyles.NoCurrentDateDefault|DateTimeStyles.AllowWhiteSpaces, out d))
            {
                return d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            return date;
        }

        static string NormalizeDateTime(string date)
        {
            DateTime d;
            if (DateTime.TryParseExact(date, "yyyy-MM-ddTHH:MM:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal | DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowWhiteSpaces, out d))
            {
                return d.ToString("yyyy-MM-ddTHH:MM:ss", CultureInfo.InvariantCulture);
            }
            if (DateTime.TryParse(date, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal | DateTimeStyles.NoCurrentDateDefault | DateTimeStyles.AllowWhiteSpaces, out d))
            {
                return d.ToString("yyyy-MM-ddTHH:MM:ss", CultureInfo.InvariantCulture);
            }
            return date;
        }

        static bool NormalizeBool(string b)
        {
            return (b.Length >= 1 && (b[0] == '1' || b[0] == 'T' || b[0] == 't' || b[0] == 'Y' || b[0] == 'y'));
        }

        static string CompileAccessibilityCodes(XPathNavigator nav)
        {
            StringBuilder builder = new StringBuilder();
            XPathNodeIterator nodes = nav.Select(sXP_AccessibilityCodes);
            {
                while (nodes.MoveNext())
                {
                    if (builder.Length > 0) builder.Append(';');
                    builder.Append(nodes.Current.ToString());
                    builder.Append(":UYU");
                }
            }
            return builder.ToString();
        }

        static int CountItemsOnPage(XPathNavigator nav, string pageNumber)
        {
            string xpath = string.Format("count(/TDSReport/Opportunity/Item[@pageNumber='{0}'])", pageNumber);
            int? count = nav.Evaluate(xpath) as int?;
            return (count != null) ? (int)count : 1; // Use 1 as default to prevent divide by zero errors.
        }

        static string CalcResponseDuration(XPathNavigator navDoc, XPathNavigator navItem)
        {
            string pageNum = navItem.Eval(sXp_PageNumber);
            string pageTime = navItem.Eval(sXp_PageTime);
            long milliseconds;
            if (!string.IsNullOrEmpty(pageNum) &&
                long.TryParse(pageTime, NumberStyles.AllowDecimalPoint|NumberStyles.AllowLeadingWhite|NumberStyles.AllowTrailingWhite, CultureInfo.InvariantCulture, out milliseconds))
            {
                int itemCount = CountItemsOnPage(navDoc, pageNum);
                if (itemCount < 1) itemCount = 1;
                return ((milliseconds / 1000.0) / (long)itemCount).ToString("F3", CultureInfo.InvariantCulture);
            }
            else
            {
                return "0";
            }
        }

        static void ProcessScoresDw(XPathNavigator nav, XPathExpression xp_ScaleScore, XPathExpression xp_StdErr,
            XPathExpression xp_PerfLvl, string[] fields, int index)
        {
            // We round fractional scale scores down to the lower whole number. This is because
            // the performance level has already been calculated on the fractional number. If
            // we rounded to the nearest whole number then half the time the number would be
            // rounded up. And if the fractional number was just below the performance level
            // cut score, a round up could show a whole number score that doesn't correspond
            // to the performance level.

            string scaleScore = Round(nav.Eval(xp_ScaleScore));
            string stdErr = Round(nav.Eval(xp_StdErr));
            string perfLvl = nav.Eval(xp_PerfLvl);

            fields[index] = scaleScore;
            fields[index + 3] = perfLvl;

            int scaleScoreN;
            int stdErrN;
            if (int.TryParse(scaleScore, out scaleScoreN) && int.TryParse(stdErr, out stdErrN))
            {
                // MinimumValue
                fields[index + 1] = (scaleScoreN - stdErrN).ToString("d", System.Globalization.CultureInfo.InvariantCulture);
                // MaximumValue
                fields[index + 2] = (scaleScoreN + stdErrN).ToString("d", System.Globalization.CultureInfo.InvariantCulture);
            }
            else
            {
                fields[index + 1] = string.Empty;
                fields[index + 2] = string.Empty;
            }
        }

        static void ProcessScoresAll(XPathNavigator nav, XPathExpression xp_ScaleScore, XPathExpression xp_StdErr,
            XPathExpression xp_PerfLvl, string[] fields, int index)
        {
            string scaleScore = nav.Eval(xp_ScaleScore);
            string stdErr = nav.Eval(xp_StdErr);
            string perfLvl = nav.Eval(xp_PerfLvl);

            fields[index] = scaleScore;
            fields[index + 1] = stdErr;
            fields[index + 2] = perfLvl;
        }

        // Take the floor of a decimal number string
        static string Floor(string number)
        {
            int point = number.IndexOf('.');
            return (point > 1) ? number.Substring(0, point) : number;
        }

        // Round a decimal number string
        static string Round(string number)
        {
            decimal v;
            if (decimal.TryParse(number, out v))
            {
                return Math.Round(v).ToString("F0", CultureInfo.InvariantCulture);
            }
            return number;
        }

        static string ByteArrayToHexString(byte[] inputArray)
        {
            if (inputArray == null) return null;
            StringBuilder o = new StringBuilder();
            for (int i = 0; i < inputArray.Length; i++) o.Append(inputArray[i].ToString("X2"));
            return o.ToString();
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
            if (m_sWriter != null)
            {
#if DEBUG
                if (!disposing) Debug.Fail("Failed to dispose CsvWriter");
#endif
                m_sWriter.Dispose();
                m_sWriter = null;
            }
            if (m_iWriter != null)
            {
                m_iWriter.Dispose();
                m_iWriter = null;
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

        public static string EvalFirstMatch(this XPathNavigator nav, params XPathExpression[] expressions)
        {
            foreach (XPathExpression exp in expressions)
            {
                if (exp.ReturnType == XPathResultType.NodeSet)
                {
                    XPathNodeIterator nodes = nav.Select(exp);
                    if (nodes.MoveNext())
                    {
                        return nodes.Current.ToString();
                    }
                }
                else
                {
                    string value = nav.Evaluate(exp).ToString();
                    if (!string.IsNullOrEmpty(value)) return value;
                }
            }
            return string.Empty;
        }

    }
}
