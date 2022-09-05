Public Class BGConstant

#Region "Constant"
    Public Const P_PIC_ALL As String = "0000"

    '// Constant Variables
    Public Const P_BUDGET_TYPE_EXPENSE As String = "E"
    Public Const P_BUDGET_TYPE_ASSET As String = "A"
    Public Const P_EXPENSE_TYPE_LABOR As String = "Labor Expense"
    Public Const P_EXPENSE_TYPE_VARIABLE As String = "Variable Expense"
    Public Const P_EXPENSE_TYPE_FIXED As String = "Fixed Expense"
    Public Const P_FC_COST As String = "Manufacturing Cost"
    Public Const P_ADMIN_COST As String = "Administration Cost"


    '// (+) Budget Compare Report
    Public Const P_RPT_DETAIL_PIC As String = "Detail by Person in Charge"
    Public Const P_RPT_SUMMARY_PIC As String = "Summary by Person in Charge"
    Public Const P_RPT_DETAIL_ACCOUNT As String = "Detail by Account No."
    Public Const P_RPT_SUMMARY_ACCOUNT As String = "Summary by Account No."
    Public Const P_RPT_INVESTMENT As String = "Summary by Investment"
    '// (-)

#End Region

#Region "Enumeration"

    '// User Level
    Public Enum enumUserLevel As Integer
        SystemAdministrator = 0
        AccountUser = 1
        ManagingDirector = 2
        AdminSaleDirector = 3
        GeneralManager = 4
        Manager = 5
        NormalUser = 6
    End Enum

    '// Budget Period Type
    Public Enum enumPeriodType As Integer
        OriginalBudget = 1
        EstimateBudget = 2
        ForecastBudget = 3
        EstimateBudget2 = 4
        ForecastBudget2 = 5
        OriginalBudget3 = 6
        EstimateBudget3 = 7
        ForecastBudget3 = 8
        ForecastBudget4 = 9
        MBPBudget = 10
        BudgetCompareVer10 = 11 '1st Half
        BudgetCompareVer20 = 12 '2nd Half
    End Enum

    '// Cost Type
    Public Enum enumCostType As Integer
        FixedCost = 1
        VariableCost = 2
    End Enum

    '// Cost
    Public Enum enumCost As Integer
        FC = 1
        ADMIN = 2
    End Enum

    '// Budget Journal Status
    Public Enum enumBudgetStatus As Integer
        NewRecord = 1
        Submit = 2
        Approve = 3
        Adjust = 4
        Authorize1 = 5
        Authorize2 = 6
    End Enum

    '// Menu Node
    Public Enum enumMenuNode As Integer
        Root
        Home
        Information
        BudgetJournal
        ViewBudget
        InputBudget
        ApproveBudget
        EditBudget
        AccountTools
        OpenNewPeriod
        ClosePeriod
        ReOpenPeriod
        ImportExportFileToSAP
        MasterData
        UserMaster
        BudgetOrderMaster
        AccountMaster
        DepartmentMaster
        PersonInChargeMaster
        CostTransferMaster
        ChangePassword
        ViewLog
        BudgetRevisionManagement
    End Enum

    '// Operation Code
    Public Enum enumOperationCd As Integer
        ViewBudget = 1
        InputBudget = 2
        ApproveBudget = 3
        AdjustBudget = 4
        Authorize1 = 5
        Authorize2 = 6
        SubmitBudget = 7
        RejectBudget1 = 8
        RejectBudget2 = 9
        UpRevision = 10
        DelRevision = 11
        RejectBudget3 = 12
        OpenNewPeriod = 13
        ClosePeriod = 14
        ReopenPeriod = 15
        ImportData = 16
        ExportData = 17
        EditInformation = 18
        EditMasterData = 19
        EditUserMaster = 20
        EditBudgetOrderMaster = 21
        EditAccountMaster = 22
        EditDepartmentMaster = 23
        EditPersonInChargeMaster = 24
        EditTransferCostMaster = 25
        EditBudgetAdjustMaster = 26
        EditAssetGroupMaster = 27
        EditChildPicMaster = 28
        EditReopenAccountMaster = 29
        EditAssetCategoryMaster = 30
        EditAssetProjectMaster = 31
        EditViewBudgetPeriod = 32
        ReInput = 33
        AdjustBudgetDirectInput = 34
        ReInputByOrder = 35
    End Enum

    '// Permission Code
    Public Enum enumPermissionCd As Integer
        Entry
        Submit
        Approve
        Adjust
        Auth1
        Auth2
        Import
        Export
        Master
        System
        View
        DirectInput
    End Enum

    '// Expense Type
    Public Enum enumExpenseType As Integer
        LaborExpense = 1
        VariableExpense = 2
        FixedExpense = 3
    End Enum

    '// Upload Data Type
    Public Enum enumUploadDataType As Integer
        BudgetData = 1
        ActualData = 2
        MTPData = 3
    End Enum

    '// Transfer Type
    Public Enum enumTransferType
        FCtoADMIN = 1
        ADMINtoFC = 2
    End Enum

    ' ''// Asset Project
    ''Public Enum enumAssetProject
    ''    TBR = 1
    ''    ORR = 2
    ''    PCT = 3
    ''End Enum

    ' ''// Asset Category
    ''Public Enum enumAssetCategory
    ''    LandAndPlantConstruction = 1
    ''    Machinery = 2
    ''    ProductionEquipment = 3
    ''    OfficeEquipment = 4
    ''    WorkingBudget = 5
    ''End Enum

#End Region

End Class
