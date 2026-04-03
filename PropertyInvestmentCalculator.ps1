[CmdletBinding()]
param()

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $argList = @(
        '-NoProfile'
        '-STA'
        '-ExecutionPolicy', 'Bypass'
        '-File', ('"{0}"' -f $PSCommandPath)
    )
    Start-Process -FilePath 'powershell' -ArgumentList $argList | Out-Null
    exit
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

function Get-DecimalValue {
    param(
        [string]$Text,
        [double]$Default = [double]::NaN
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Default
    }

    $clean = $Text.Trim() -replace '[^0-9,.-]', ''
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $Default
    }

    $value = 0.0
    if ([double]::TryParse($clean, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
        return $value
    }

    $normalized = $clean -replace ',', '.'
    if ([double]::TryParse($normalized, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value)) {
        return $value
    }

    return $Default
}

function Format-Currency {
    param([double]$Value)
    return ('{0:C2}' -f $Value)
}

function Format-Percent {
    param([double]$Value)
    return ('{0:N2}%' -f ($Value * 100))
}

function Safe-Divide {
    param(
        [double]$Numerator,
        [double]$Denominator
    )

    if ([Math]::Abs($Denominator) -lt 0.0000001) {
        return 0
    }

    return $Numerator / $Denominator
}

function Get-EnglandNiTax {
    param(
        [double]$Price,
        [bool]$IsAdditional
    )

    $bands = @(
        @{ Limit = 125000; Rate = 0.00 },
        @{ Limit = 250000; Rate = 0.02 },
        @{ Limit = 925000; Rate = 0.05 },
        @{ Limit = 1500000; Rate = 0.10 },
        @{ Limit = [double]::PositiveInfinity; Rate = 0.12 }
    )

    $extra = if ($IsAdditional) { 0.05 } else { 0.00 }
    $tax = 0.0
    $previous = 0.0

    foreach ($band in $bands) {
        $upper = [double]$band.Limit
        $taxable = [Math]::Min($Price, $upper) - $previous
        if ($taxable -gt 0) {
            $tax += $taxable * ($band.Rate + $extra)
        }
        if ($Price -le $upper) {
            break
        }
        $previous = $upper
    }

    return $tax
}

function Get-WalesTax {
    param(
        [double]$Price,
        [bool]$IsAdditional
    )

    if ($IsAdditional) {
        $bands = @(
            @{ Limit = 180000; Rate = 0.05 },
            @{ Limit = 250000; Rate = 0.085 },
            @{ Limit = 400000; Rate = 0.10 },
            @{ Limit = 750000; Rate = 0.125 },
            @{ Limit = 1500000; Rate = 0.15 },
            @{ Limit = [double]::PositiveInfinity; Rate = 0.17 }
        )
    }
    else {
        $bands = @(
            @{ Limit = 225000; Rate = 0.00 },
            @{ Limit = 400000; Rate = 0.06 },
            @{ Limit = 750000; Rate = 0.075 },
            @{ Limit = 1500000; Rate = 0.10 },
            @{ Limit = [double]::PositiveInfinity; Rate = 0.12 }
        )
    }

    $tax = 0.0
    $previous = 0.0

    foreach ($band in $bands) {
        $upper = [double]$band.Limit
        $taxable = [Math]::Min($Price, $upper) - $previous
        if ($taxable -gt 0) {
            $tax += $taxable * $band.Rate
        }
        if ($Price -le $upper) {
            break
        }
        $previous = $upper
    }

    return $tax
}

function Get-PropertyTax {
    param(
        [double]$Price,
        [string]$Region,
        [bool]$IsAdditional,
        [double]$ManualStampDuty
    )

    switch ($Region) {
        'England / NI' { return Get-EnglandNiTax -Price $Price -IsAdditional $IsAdditional }
        'Wales' { return Get-WalesTax -Price $Price -IsAdditional $IsAdditional }
        default { return $ManualStampDuty }
    }
}

function Get-BtlAnalysis {
    param([hashtable]$Data)

    $stampDuty = Get-PropertyTax -Price $Data.PurchasePrice -Region $Data.Region -IsAdditional $Data.IsAdditional -ManualStampDuty $Data.ManualStampDuty
    $depositAmount = $Data.PurchasePrice * $Data.DepositPct
    $loanAmount = $Data.PurchasePrice - $depositAmount
    $monthlyMortgage = $loanAmount * $Data.InterestRate / 12
    $financeCost = $monthlyMortgage * $Data.FinanceMonths
    $holdingCost = $Data.VacantMonthlyCosts * $Data.VacantMonths
    $buyingCosts = $Data.LegalFees + $Data.BrokerFees + $Data.SurveyFees + $Data.SourcingFees + $Data.OtherBuyingCosts + $stampDuty
    $upfrontCash = $depositAmount + $buyingCosts + $Data.RefurbCost
    $totalCashIn = $upfrontCash + $financeCost + $holdingCost
    $monthlyCashflow = $Data.MonthlyRent - $Data.MonthlyOperatingCosts - $monthlyMortgage
    $annualCashflow = $monthlyCashflow * 12
    $grossYield = Safe-Divide -Numerator ($Data.MonthlyRent * 12) -Denominator $Data.PurchasePrice
    $cashOnCash = Safe-Divide -Numerator $annualCashflow -Denominator $totalCashIn
    $breakEvenOccupancy = Safe-Divide -Numerator ($Data.MonthlyOperatingCosts + $monthlyMortgage) -Denominator $Data.MonthlyRent

    return [ordered]@{
        StampDuty = $stampDuty
        BuyingCosts = $buyingCosts
        DepositAmount = $depositAmount
        LoanAmount = $loanAmount
        MonthlyMortgage = $monthlyMortgage
        FinanceCost = $financeCost
        HoldingCost = $holdingCost
        UpfrontCash = $upfrontCash
        TotalCashIn = $totalCashIn
        MonthlyCashflow = $monthlyCashflow
        AnnualCashflow = $annualCashflow
        GrossYield = $grossYield
        CashOnCash = $cashOnCash
        BreakEvenOccupancy = $breakEvenOccupancy
    }
}

function Get-BrrrAnalysis {
    param([hashtable]$Data)

    $stampDuty = Get-PropertyTax -Price $Data.PurchasePrice -Region $Data.Region -IsAdditional $Data.IsAdditional -ManualStampDuty $Data.ManualStampDuty
    $purchaseDeposit = $Data.PurchasePrice * $Data.PurchaseDepositPct
    $purchaseLoan = $Data.PurchasePrice - $purchaseDeposit
    $monthlyBridgeCost = $purchaseLoan * $Data.PurchaseInterestRate / 12
    $bridgeCostTotal = $monthlyBridgeCost * $Data.FinanceMonths
    $vacantHoldingTotal = $Data.VacantMonthlyCosts * $Data.VacantMonths
    $buyingCosts = $Data.LegalFees + $Data.BrokerFees + $Data.SurveyFees + $Data.SourcingFees + $Data.OtherBuyingCosts + $stampDuty
    $cashBeforeRefi = $purchaseDeposit + $buyingCosts + $Data.RefurbCost + $bridgeCostTotal + $vacantHoldingTotal
    $refinanceMortgage = $Data.EndValue * (1 - $Data.RefinanceDepositPct)
    $moneyTakenOut = [Math]::Max(0, $refinanceMortgage - $purchaseLoan)
    $cashLeftIn = $cashBeforeRefi + $Data.RefinanceCosts - $moneyTakenOut
    $monthlyRefiMortgage = $refinanceMortgage * $Data.RefinanceInterestRate / 12
    $monthlyCashflow = $Data.MonthlyRent - $Data.MonthlyOperatingCosts - $monthlyRefiMortgage
    $annualCashflow = $monthlyCashflow * 12
    $grossYieldOnPurchase = Safe-Divide -Numerator ($Data.MonthlyRent * 12) -Denominator $Data.PurchasePrice
    $grossYieldOnEndValue = Safe-Divide -Numerator ($Data.MonthlyRent * 12) -Denominator $Data.EndValue
    $roce = Safe-Divide -Numerator $annualCashflow -Denominator $cashLeftIn
    $equityCreated = $Data.EndValue - ($Data.PurchasePrice + $Data.RefurbCost + $buyingCosts)

    return [ordered]@{
        StampDuty = $stampDuty
        BuyingCosts = $buyingCosts
        PurchaseDeposit = $purchaseDeposit
        PurchaseLoan = $purchaseLoan
        MonthlyBridgeCost = $monthlyBridgeCost
        BridgeCostTotal = $bridgeCostTotal
        VacantHoldingTotal = $vacantHoldingTotal
        CashBeforeRefi = $cashBeforeRefi
        RefinanceMortgage = $refinanceMortgage
        MoneyTakenOut = $moneyTakenOut
        CashLeftIn = $cashLeftIn
        MonthlyRefiMortgage = $monthlyRefiMortgage
        MonthlyCashflow = $monthlyCashflow
        AnnualCashflow = $annualCashflow
        GrossYieldOnPurchase = $grossYieldOnPurchase
        GrossYieldOnEndValue = $grossYieldOnEndValue
        ROCE = $roce
        EquityCreated = $equityCreated
    }
}

function Get-HmoAnalysis {
    param([hashtable]$Data)

    $stampDuty = Get-PropertyTax -Price $Data.PurchasePrice -Region $Data.Region -IsAdditional $Data.IsAdditional -ManualStampDuty $Data.ManualStampDuty
    $upfrontTotal = $Data.PurchasePrice + $Data.SourcingFee + $Data.DevelopmentCost + $Data.ProjectManagement + $Data.FurniturePurchased + $Data.PurchaseCosts + $stampDuty
    $grossMonthlyRent = [double](($Data.RoomRents | Measure-Object -Sum).Sum)
    $managementFee = $grossMonthlyRent * $Data.ManagementFeePct
    $mortgageAmount = $Data.RefinanceValue * $Data.LtvPct
    $monthlyMortgage = $mortgageAmount * $Data.MortgageRate / 12
    $monthlyInvestorCost = $Data.InvestorFunds * $Data.InvestorRate / 12
    $totalMonthlyCosts = $monthlyMortgage + $Data.MonthlyOperatingCosts + $managementFee + $Data.MonthlyUtilities + $Data.MonthlyFurnitureLease + $monthlyInvestorCost
    $monthlyCashflow = $grossMonthlyRent - $totalMonthlyCosts
    $annualCashflow = $monthlyCashflow * 12
    $cashLeftIn = $upfrontTotal - $mortgageAmount - $Data.InvestorFunds
    $roi = Safe-Divide -Numerator $annualCashflow -Denominator $cashLeftIn
    $grossYield = Safe-Divide -Numerator ($grossMonthlyRent * 12) -Denominator $Data.PurchasePrice
    $commercialValue = Safe-Divide -Numerator ($grossMonthlyRent * 12) -Denominator $Data.YieldTarget
    $equityOrProfit = $commercialValue - $upfrontTotal

    return [ordered]@{
        StampDuty = $stampDuty
        UpfrontTotal = $upfrontTotal
        GrossMonthlyRent = $grossMonthlyRent
        ManagementFee = $managementFee
        MortgageAmount = $mortgageAmount
        MonthlyMortgage = $monthlyMortgage
        MonthlyInvestorCost = $monthlyInvestorCost
        TotalMonthlyCosts = $totalMonthlyCosts
        MonthlyCashflow = $monthlyCashflow
        AnnualCashflow = $annualCashflow
        CashLeftIn = $cashLeftIn
        ROI = $roi
        GrossYield = $grossYield
        CommercialValue = $commercialValue
        EquityOrProfit = $equityOrProfit
    }
}

function Get-DealStatus {
    param(
        [string]$Strategy,
        [hashtable]$Result
    )

    switch ($Strategy) {
        'BTL' {
            if ($Result.MonthlyCashflow -gt 0 -and $Result.GrossYield -ge 0.08 -and $Result.CashOnCash -ge 0.12) {
                return @{ Text = 'Good deal'; Color = '#81C784'; Note = 'Positive cashflow, strong gross yield and healthy cash-on-cash return.' }
            }
            if ($Result.MonthlyCashflow -gt 0 -and ($Result.GrossYield -ge 0.06 -or $Result.CashOnCash -ge 0.08)) {
                return @{ Text = 'Needs review'; Color = '#FFD54F'; Note = 'Deal is workable, but margin is thinner and should be stress-tested.' }
            }
            return @{ Text = 'Weak deal'; Color = '#E57373'; Note = 'Cashflow or return is too weak for a comfortable buy-to-let margin.' }
        }
        'BRRR' {
            if ($Result.MonthlyCashflow -gt 0 -and $Result.ROCE -ge 0.20 -and $Result.CashLeftIn -le ($Result.CashBeforeRefi * 0.60)) {
                return @{ Text = 'Good deal'; Color = '#81C784'; Note = 'Refinance leaves limited cash in the deal while maintaining strong ROCE.' }
            }
            if ($Result.MonthlyCashflow -gt 0 -and $Result.ROCE -ge 0.10) {
                return @{ Text = 'Needs review'; Color = '#FFD54F'; Note = 'The BRRR works on paper, but the refinance outcome needs checking.' }
            }
            return @{ Text = 'Weak deal'; Color = '#E57373'; Note = 'Too much cash stays in, or the refinance cashflow is too weak.' }
        }
        default {
            if ($Result.MonthlyCashflow -gt 0 -and $Result.ROI -ge 0.15 -and $Result.GrossYield -ge 0.10) {
                return @{ Text = 'Good deal'; Color = '#81C784'; Note = 'Strong HMO cashflow and ROI with good rent-to-value performance.' }
            }
            if ($Result.MonthlyCashflow -gt 0 -and $Result.ROI -ge 0.08) {
                return @{ Text = 'Needs review'; Color = '#FFD54F'; Note = 'The HMO can work, but rent, bills and management need tighter checks.' }
            }
            return @{ Text = 'Weak deal'; Color = '#E57373'; Note = 'Cashflow or ROI is too weak for the complexity and risk of an HMO.' }
        }
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $prefix = "[$timestamp] [$Level]"
    $script:ui.LogBox.AppendText("$prefix $Message`r`n")
    $script:ui.LogBox.ScrollToEnd()
}

function Set-BusyLabel {
    param([string]$Text)
    $script:ui.BusyLabel.Text = $Text
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="UK Property Investment Calculator"
        Height="940"
        Width="1500"
        MinHeight="860"
        MinWidth="1320"
        WindowStartupLocation="CenterScreen"
        Background="#141B24"
        FontFamily="Arial">
    <Window.Resources>
        <SolidColorBrush x:Key="WindowBrush" Color="#141B24"/>
        <SolidColorBrush x:Key="SidebarBrush" Color="#101720"/>
        <SolidColorBrush x:Key="CardBrush" Color="#1C2430"/>
        <SolidColorBrush x:Key="ActiveBrush" Color="#12293D"/>
        <SolidColorBrush x:Key="CardHighlightBrush" Color="#263246"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#00A6C8"/>
        <SolidColorBrush x:Key="AccentSoftBrush" Color="#19D3F3"/>
        <SolidColorBrush x:Key="ForegroundBrush" Color="#C5D1DE"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#94A3B8"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#0D567A"/>
        <SolidColorBrush x:Key="RowBorderBrush" Color="#243446"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#81C784"/>
        <SolidColorBrush x:Key="WarningBrush" Color="#FFD54F"/>
        <SolidColorBrush x:Key="DangerBrush" Color="#E57373"/>
        <LinearGradientBrush x:Key="HeroBrush" StartPoint="0,0" EndPoint="1,1">
            <GradientStop Color="#1B2634" Offset="0"/>
            <GradientStop Color="#12293D" Offset="0.65"/>
            <GradientStop Color="#0D567A" Offset="1"/>
        </LinearGradientBrush>
        <LinearGradientBrush x:Key="PrimaryGradientBrush" StartPoint="0,0" EndPoint="1,0">
            <GradientStop Color="#00A6C8" Offset="0"/>
            <GradientStop Color="#19D3F3" Offset="1"/>
        </LinearGradientBrush>

        <Style x:Key="BaseButtonStyle" TargetType="Button">
            <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
            <Setter Property="Background" Value="{StaticResource CardBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="Margin" Value="0,0,8,8"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Padding" Value="12,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="9" SnapsToDevicePixels="True">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="{TemplateBinding Padding}"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="PrimaryButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="{StaticResource PrimaryGradientBrush}"/>
            <Setter Property="BorderBrush" Value="Transparent"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>

        <Style x:Key="SecondaryButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Background" Value="#202B39"/>
        </Style>

        <Style x:Key="NavButtonStyle" TargetType="Button" BasedOn="{StaticResource BaseButtonStyle}">
            <Setter Property="Width" Value="210"/>
            <Setter Property="Height" Value="46"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Padding" Value="16,0,0,0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Background" Value="#1E2836"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
        </Style>

        <Style x:Key="InputTextBoxStyle" TargetType="TextBox">
            <Setter Property="Height" Value="30"/>
            <Setter Property="Background" Value="{StaticResource ActiveBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="MinWidth" Value="140"/>
        </Style>

        <Style x:Key="InputComboStyle" TargetType="ComboBox">
            <Setter Property="Height" Value="30"/>
            <Setter Property="Background" Value="{StaticResource ActiveBrush}"/>
            <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="MinWidth" Value="170"/>
        </Style>

        <Style x:Key="InputCheckStyle" TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
            <Setter Property="Margin" Value="0,2,0,10"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>

        <Style x:Key="FieldLabelStyle" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource ForegroundBrush}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="TextWrapping" Value="Wrap"/>
            <Setter Property="Margin" Value="0,3,14,10"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>

        <Style x:Key="BadgeStyle" TargetType="Border">
            <Setter Property="Background" Value="#0D567A"/>
            <Setter Property="BorderBrush" Value="#00A6C8"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="CornerRadius" Value="12"/>
            <Setter Property="Padding" Value="10,4"/>
        </Style>
    </Window.Resources>

    <Grid Background="{StaticResource WindowBrush}">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="330"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="{StaticResource SidebarBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="0,0,1,0">
            <Grid Margin="16">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" Background="{StaticResource HeroBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" CornerRadius="16" Padding="16" Margin="0,0,0,14">
                    <StackPanel>
                        <TextBlock Text="Property Calculator" Foreground="White" FontSize="26" FontWeight="Bold" TextWrapping="Wrap"/>
                        <TextBlock Text="BTL / BRRR / HMO analysis for UK deals" Foreground="#D6E7F2" Margin="0,6,0,0" TextWrapping="Wrap"/>
                        <TextBlock Text="Screen deals faster, pitch them better, and export investor-ready reports." Foreground="#A9C8D8" Margin="0,10,0,0" TextWrapping="Wrap"/>
                        <Border Style="{StaticResource BadgeStyle}" Margin="0,12,0,0" HorizontalAlignment="Stretch" Background="#0F3850">
                            <TextBlock Text="Powered by marcinsakowski.com" Foreground="White" FontWeight="SemiBold" TextWrapping="Wrap"/>
                        </Border>
                    </StackPanel>
                </Border>

                <StackPanel Grid.Row="1" Margin="0,0,0,12">
                    <Button Name="NavBTL" Content="BTL" Style="{StaticResource NavButtonStyle}"/>
                    <Button Name="NavBRRR" Content="BRRR" Style="{StaticResource NavButtonStyle}"/>
                    <Button Name="NavHMO" Content="HMO" Style="{StaticResource NavButtonStyle}"/>
                </StackPanel>

                <Border Grid.Row="2" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="10" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Quick Notes" Foreground="{StaticResource ForegroundBrush}" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
                        <TextBlock Text="Use the calculator as a deal screening tool, not as legal, tax or mortgage advice." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,0,0,8"/>
                        <TextBlock Text="Auto tax is included for England / NI and Wales. Use Manual for Scotland or custom cases." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                    </StackPanel>
                </Border>

                <Border Grid.Row="3" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="10" Margin="0,0,0,12">
                    <StackPanel>
                        <TextBlock Text="Share-Ready" Foreground="{StaticResource ForegroundBrush}" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
                        <TextBlock Text="Brand it and share it with clients, partners or other investors." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,0,0,10"/>
                        <Button Name="WebsiteButton" Content="marcinsakowski.com" Style="{StaticResource SecondaryButtonStyle}" HorizontalAlignment="Stretch"/>
                        <Button Name="BookReviewButton" Content="Book a deal review" Style="{StaticResource PrimaryButtonStyle}" HorizontalAlignment="Stretch"/>
                        <TextBlock Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,6,0,0">
                            <Run Text="Website: "/>
                            <Hyperlink Name="WebsiteLink" NavigateUri="https://marcinsakowski.com">https://marcinsakowski.com</Hyperlink>
                        </TextBlock>
                    </StackPanel>
                </Border>

                <Border Grid.Row="4" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="10">
                    <StackPanel>
                        <TextBlock Text="Lead Capture" Foreground="{StaticResource ForegroundBrush}" FontSize="14" FontWeight="SemiBold" Margin="0,0,0,8"/>
                        <TextBox Name="LeadName" Style="{StaticResource InputTextBoxStyle}" Text="Client / investor name"/>
                        <TextBox Name="LeadEmail" Style="{StaticResource InputTextBoxStyle}" Text="Email address"/>
                        <TextBox Name="LeadPhone" Style="{StaticResource InputTextBoxStyle}" Text="Phone number"/>
                        <TextBox Name="LeadSource" Style="{StaticResource InputTextBoxStyle}" Text="Lead source / notes"/>
                    </StackPanel>
                </Border>

                <TextBlock Grid.Row="5" Text="Built for property deal analysis" Foreground="{StaticResource MutedBrush}" Margin="0,12,0,0"/>
            </Grid>
        </Border>

        <Grid Grid.Column="1" Margin="18">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="230"/>
            </Grid.RowDefinitions>

            <Border Grid.Row="0" Background="{StaticResource HeroBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" CornerRadius="18" Padding="18" Margin="0,0,0,12">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel>
                        <TextBlock Text="UK Property Investment Calculator" Foreground="White" FontSize="26" FontWeight="Bold"/>
                        <TextBlock Name="HeaderSubtitle" Text="Choose a strategy from the left and analyse the deal." Foreground="#D6E7F2" Margin="0,6,0,0" TextWrapping="Wrap"/>
                        <TextBlock Text="Built to look credible in front of clients, JV partners and investors." Foreground="#A9C8D8" Margin="0,10,0,0" TextWrapping="Wrap"/>
                    </StackPanel>
                    <WrapPanel Grid.Column="1">
                        <Border Style="{StaticResource BadgeStyle}" Background="#0F3850" Margin="0,0,6,0"><TextBlock Text="Responsive layout" Foreground="White"/></Border>
                        <Border Style="{StaticResource BadgeStyle}" Background="#0F3850" Margin="0,0,6,0"><TextBlock Text="Validation" Foreground="White"/></Border>
                        <Border Style="{StaticResource BadgeStyle}" Background="#0F3850"><TextBlock Text="Deal scoring" Foreground="White"/></Border>
                    </WrapPanel>
                </Grid>
            </Border>

            <Grid Grid.Row="1">
                <Grid Name="BTLPage">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="10" Margin="0,0,0,10">
                        <WrapPanel>
                            <Button Name="BTLCalculate" Content="Calculate BTL" Style="{StaticResource PrimaryButtonStyle}"/>
                            <Button Name="BTLReset" Content="Reset defaults" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BTLSave" Content="Save deal" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BTLLoad" Content="Load deal" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BTLCopy" Content="Copy summary" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BTLExport" Content="Export CSV" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BTLExportXlsx" Content="Export XLSX" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BTLExportPdf" Content="Export PDF" Style="{StaticResource SecondaryButtonStyle}"/>
                            <TextBlock Text="Classic buy-to-let analysis with purchase, finance, rent and cashflow." Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="8,0,0,8" TextWrapping="Wrap"/>
                        </WrapPanel>
                    </Border>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="430"/>
                            <ColumnDefinition Width="14"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
                            <StackPanel Name="BTLFormHost"/>
                        </ScrollViewer>
                        <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
                            <StackPanel Name="BTLResultHost"/>
                        </ScrollViewer>
                    </Grid>
                </Grid>

                <Grid Name="BRRRPage" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="10" Margin="0,0,0,10">
                        <WrapPanel>
                            <Button Name="BRRRCalculate" Content="Calculate BRRR" Style="{StaticResource PrimaryButtonStyle}"/>
                            <Button Name="BRRRReset" Content="Reset defaults" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BRRRSave" Content="Save deal" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BRRRLoad" Content="Load deal" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BRRRCopy" Content="Copy summary" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BRRRExport" Content="Export CSV" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BRRRExportXlsx" Content="Export XLSX" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="BRRRExportPdf" Content="Export PDF" Style="{StaticResource SecondaryButtonStyle}"/>
                            <TextBlock Text="Buy, refurbish, refinance and rent analysis with cash left in deal." Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="8,0,0,8" TextWrapping="Wrap"/>
                        </WrapPanel>
                    </Border>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="430"/>
                            <ColumnDefinition Width="14"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
                            <StackPanel Name="BRRRFormHost"/>
                        </ScrollViewer>
                        <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
                            <StackPanel Name="BRRRResultHost"/>
                        </ScrollViewer>
                    </Grid>
                </Grid>

                <Grid Name="HMOPage" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="10" Margin="0,0,0,10">
                        <WrapPanel>
                            <Button Name="HMOCalculate" Content="Calculate HMO" Style="{StaticResource PrimaryButtonStyle}"/>
                            <Button Name="HMOReset" Content="Reset defaults" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="HMOSave" Content="Save deal" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="HMOLoad" Content="Load deal" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="HMOCopy" Content="Copy summary" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="HMOExport" Content="Export CSV" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="HMOExportXlsx" Content="Export XLSX" Style="{StaticResource SecondaryButtonStyle}"/>
                            <Button Name="HMOExportPdf" Content="Export PDF" Style="{StaticResource SecondaryButtonStyle}"/>
                            <TextBlock Text="Multi-room HMO analysis with rent-by-room, bills, management and ROI." Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="8,0,0,8" TextWrapping="Wrap"/>
                        </WrapPanel>
                    </Border>
                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="430"/>
                            <ColumnDefinition Width="14"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
                            <StackPanel Name="HMOFormHost"/>
                        </ScrollViewer>
                        <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
                            <StackPanel Name="HMOResultHost"/>
                        </ScrollViewer>
                    </Grid>
                </Grid>
            </Grid>

            <Border Grid.Row="2" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Padding="12" Margin="0,12,0,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Status" Foreground="{StaticResource ForegroundBrush}" FontSize="14" FontWeight="SemiBold"/>
                        <TextBlock Name="BusyLabel" Grid.Column="1" Text="Ready" Foreground="{StaticResource AccentBrush}" FontSize="13" FontWeight="SemiBold" Margin="12,0,0,0"/>
                    </Grid>
                    <TextBox Name="LogBox" Grid.Row="1" Background="#12293D" Foreground="{StaticResource ForegroundBrush}" BorderBrush="{StaticResource BorderBrush}" AcceptsReturn="True" IsReadOnly="True" VerticalScrollBarVisibility="Visible" HorizontalScrollBarVisibility="Visible" TextWrapping="Wrap" FontFamily="Consolas" FontSize="13"/>
                </Grid>
            </Border>
        </Grid>
    </Grid>
</Window>
"@

try {
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
}
catch {
    [System.Windows.MessageBox]::Show(
        "Failed to load the calculator UI.`n`n$($_.Exception.Message)",
        'UK Property Investment Calculator',
        'OK',
        'Error'
    ) | Out-Null
    throw
}

$script:ui = [ordered]@{
    Window = $window
    NavBTL = $window.FindName('NavBTL')
    NavBRRR = $window.FindName('NavBRRR')
    NavHMO = $window.FindName('NavHMO')
    BTLPage = $window.FindName('BTLPage')
    BRRRPage = $window.FindName('BRRRPage')
    HMOPage = $window.FindName('HMOPage')
    BTLFormHost = $window.FindName('BTLFormHost')
    BRRRFormHost = $window.FindName('BRRRFormHost')
    HMOFormHost = $window.FindName('HMOFormHost')
    BTLResultHost = $window.FindName('BTLResultHost')
    BRRRResultHost = $window.FindName('BRRRResultHost')
    HMOResultHost = $window.FindName('HMOResultHost')
    BTLCalculate = $window.FindName('BTLCalculate')
    BRRRCalculate = $window.FindName('BRRRCalculate')
    HMOCalculate = $window.FindName('HMOCalculate')
    BTLReset = $window.FindName('BTLReset')
    BRRRReset = $window.FindName('BRRRReset')
    HMOReset = $window.FindName('HMOReset')
    BTLExport = $window.FindName('BTLExport')
    BRRRExport = $window.FindName('BRRRExport')
    HMOExport = $window.FindName('HMOExport')
    BTLSave = $window.FindName('BTLSave')
    BRRRSave = $window.FindName('BRRRSave')
    HMOSave = $window.FindName('HMOSave')
    BTLLoad = $window.FindName('BTLLoad')
    BRRRLoad = $window.FindName('BRRRLoad')
    HMOLoad = $window.FindName('HMOLoad')
    BTLCopy = $window.FindName('BTLCopy')
    BRRRCopy = $window.FindName('BRRRCopy')
    HMOCopy = $window.FindName('HMOCopy')
    BTLExportXlsx = $window.FindName('BTLExportXlsx')
    BRRRExportXlsx = $window.FindName('BRRRExportXlsx')
    HMOExportXlsx = $window.FindName('HMOExportXlsx')
    BTLExportPdf = $window.FindName('BTLExportPdf')
    BRRRExportPdf = $window.FindName('BRRRExportPdf')
    HMOExportPdf = $window.FindName('HMOExportPdf')
    HeaderSubtitle = $window.FindName('HeaderSubtitle')
    BusyLabel = $window.FindName('BusyLabel')
    LogBox = $window.FindName('LogBox')
    WebsiteButton = $window.FindName('WebsiteButton')
    WebsiteLink = $window.FindName('WebsiteLink')
    LeadName = $window.FindName('LeadName')
    LeadEmail = $window.FindName('LeadEmail')
    LeadPhone = $window.FindName('LeadPhone')
    LeadSource = $window.FindName('LeadSource')
    CardBrush = $window.FindResource('CardBrush')
    ActiveBrush = $window.FindResource('ActiveBrush')
    BorderBrush = $window.FindResource('BorderBrush')
    ForegroundBrush = $window.FindResource('ForegroundBrush')
    MutedBrush = $window.FindResource('MutedBrush')
    AccentBrush = $window.FindResource('AccentBrush')
    SuccessBrush = $window.FindResource('SuccessBrush')
    WarningBrush = $window.FindResource('WarningBrush')
    DangerBrush = $window.FindResource('DangerBrush')
    HeroBrush = $window.FindResource('HeroBrush')
    BookReviewButton = $window.FindName('BookReviewButton')
    NavButtonStyle = $window.FindResource('NavButtonStyle')
    SecondaryButtonStyle = $window.FindResource('SecondaryButtonStyle')
    PrimaryButtonStyle = $window.FindResource('PrimaryButtonStyle')
    InputTextBoxStyle = $window.FindResource('InputTextBoxStyle')
    InputComboStyle = $window.FindResource('InputComboStyle')
    InputCheckStyle = $window.FindResource('InputCheckStyle')
    FieldLabelStyle = $window.FindResource('FieldLabelStyle')
}

$script:pages = @{}

function New-CardBorder {
    param([string]$Title)

    $border = [System.Windows.Controls.Border]::new()
    $border.Background = $script:ui.CardBrush
    $border.BorderBrush = $script:ui.BorderBrush
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.Padding = [System.Windows.Thickness]::new(16)
    $border.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14)
    $border.CornerRadius = [System.Windows.CornerRadius]::new(16)
    $border.Effect = [System.Windows.Media.Effects.DropShadowEffect]@{
        BlurRadius = 18
        ShadowDepth = 0
        Opacity = 0.20
        Color = [System.Windows.Media.ColorConverter]::ConvertFromString('#05080D')
    }

    $stack = [System.Windows.Controls.StackPanel]::new()
    $titleBlock = [System.Windows.Controls.TextBlock]::new()
    $titleBlock.Text = $Title
    $titleBlock.Foreground = $script:ui.ForegroundBrush
    $titleBlock.FontSize = 17
    $titleBlock.FontWeight = 'SemiBold'
    $titleBlock.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14)
    [void]$stack.Children.Add($titleBlock)
    $border.Child = $stack

    return @{ Border = $border; Body = $stack }
}

function Add-FieldsGrid {
    param(
        [System.Windows.Controls.Panel]$Parent,
        [array]$Fields,
        [hashtable]$Store,
        [System.Windows.Window]$Window
    )

    $grid = [System.Windows.Controls.Grid]::new()
    $grid.Margin = [System.Windows.Thickness]::new(0)
    [void]$grid.ColumnDefinitions.Add(([System.Windows.Controls.ColumnDefinition]@{ Width = [System.Windows.GridLength]::new(200) }))
    [void]$grid.ColumnDefinitions.Add(([System.Windows.Controls.ColumnDefinition]@{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))

    $rowIndex = 0
    foreach ($field in $Fields) {
        $row = [System.Windows.Controls.RowDefinition]::new()
        $row.Height = [System.Windows.GridLength]::Auto
        [void]$grid.RowDefinitions.Add($row)

        if ($field.Type -eq 'check') {
            $check = [System.Windows.Controls.CheckBox]::new()
            $check.Style = $script:ui.InputCheckStyle
            $check.Content = $field.Label
            $check.IsChecked = [bool]$field.Default
            [System.Windows.Controls.Grid]::SetRow($check, $rowIndex)
            [System.Windows.Controls.Grid]::SetColumnSpan($check, 2)
            [void]$grid.Children.Add($check)
            $Store[$field.Key] = $check
            $Store.Meta[$field.Key] = $field
            $rowIndex++
            continue
        }

        $label = [System.Windows.Controls.TextBlock]::new()
        $label.Style = $script:ui.FieldLabelStyle
        $label.Text = $field.Label
        [System.Windows.Controls.Grid]::SetRow($label, $rowIndex)
        [System.Windows.Controls.Grid]::SetColumn($label, 0)
        [void]$grid.Children.Add($label)

        switch ($field.Type) {
            'combo' {
                $control = [System.Windows.Controls.ComboBox]::new()
                $control.Style = $script:ui.InputComboStyle
                $control.ItemsSource = $field.Options
                $control.SelectedItem = $field.Default
            }
            default {
                $control = [System.Windows.Controls.TextBox]::new()
                $control.Style = $script:ui.InputTextBoxStyle
                $control.Text = [string]$field.Default
            }
        }

        [System.Windows.Controls.Grid]::SetRow($control, $rowIndex)
        [System.Windows.Controls.Grid]::SetColumn($control, 1)
        [void]$grid.Children.Add($control)
        $Store[$field.Key] = $control
        $Store.Meta[$field.Key] = $field
        $rowIndex++
    }

    [void]$Parent.Children.Add($grid)
}

function Initialize-ResultPanel {
    param(
        [System.Windows.Controls.Panel]$ResultPanel,
        [string]$Strategy,
        [string]$Description
    )

    $headerCard = New-CardBorder -Title "$Strategy deal result"
    $topGrid = [System.Windows.Controls.Grid]::new()
    [void]$topGrid.ColumnDefinitions.Add(([System.Windows.Controls.ColumnDefinition]@{ Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }))
    [void]$topGrid.ColumnDefinitions.Add(([System.Windows.Controls.ColumnDefinition]@{ Width = [System.Windows.GridLength]::Auto }))

    $desc = [System.Windows.Controls.TextBlock]::new()
    $desc.Text = $Description
    $desc.Foreground = $script:ui.MutedBrush
    $desc.TextWrapping = 'Wrap'
    $desc.Margin = [System.Windows.Thickness]::new(0, 0, 12, 12)
    [System.Windows.Controls.Grid]::SetColumn($desc, 0)
    [void]$topGrid.Children.Add($desc)

    $badge = [System.Windows.Controls.Border]::new()
    $badge.CornerRadius = [System.Windows.CornerRadius]::new(14)
    $badge.Padding = [System.Windows.Thickness]::new(14, 8, 14, 8)
    $badge.Background = $script:ui.WarningBrush
    $badge.HorizontalAlignment = 'Right'
    $badge.VerticalAlignment = 'Top'
    [System.Windows.Controls.Grid]::SetColumn($badge, 1)

    $badgeText = [System.Windows.Controls.TextBlock]::new()
    $badgeText.Text = 'Not calculated yet'
    $badgeText.Foreground = [System.Windows.Media.Brushes]::Black
    $badgeText.FontWeight = 'SemiBold'
    $badgeText.FontSize = 13
    $badge.Child = $badgeText
    [void]$topGrid.Children.Add($badge)
    [void]$headerCard.Body.Children.Add($topGrid)

    $summaryGrid = [System.Windows.Controls.Primitives.UniformGrid]::new()
    $summaryGrid.Columns = 2
    $summaryGrid.Margin = [System.Windows.Thickness]::new(0, 6, 0, 12)
    [void]$headerCard.Body.Children.Add($summaryGrid)

    $note = [System.Windows.Controls.TextBlock]::new()
    $note.Foreground = $script:ui.MutedBrush
    $note.TextWrapping = 'Wrap'
    [void]$headerCard.Body.Children.Add($note)

    $execCard = New-CardBorder -Title 'Executive summary'
    $execHeadline = [System.Windows.Controls.TextBlock]::new()
    $execHeadline.Text = 'Run the calculator to generate an investor-ready summary.'
    $execHeadline.Foreground = $script:ui.ForegroundBrush
    $execHeadline.FontSize = 18
    $execHeadline.FontWeight = 'Bold'
    $execHeadline.TextWrapping = 'Wrap'
    $execHeadline.Margin = [System.Windows.Thickness]::new(0,0,0,8)
    [void]$execCard.Body.Children.Add($execHeadline)

    $execBody = [System.Windows.Controls.TextBlock]::new()
    $execBody.Foreground = $script:ui.MutedBrush
    $execBody.TextWrapping = 'Wrap'
    [void]$execCard.Body.Children.Add($execBody)

    $pitchCard = New-CardBorder -Title 'Share-ready pitch'
    $pitchBox = [System.Windows.Controls.TextBox]::new()
    $pitchBox.Background = $script:ui.ActiveBrush
    $pitchBox.Foreground = $script:ui.ForegroundBrush
    $pitchBox.BorderBrush = $script:ui.BorderBrush
    $pitchBox.BorderThickness = [System.Windows.Thickness]::new(1)
    $pitchBox.FontFamily = 'Consolas'
    $pitchBox.FontSize = 13
    $pitchBox.AcceptsReturn = $true
    $pitchBox.TextWrapping = 'Wrap'
    $pitchBox.IsReadOnly = $true
    $pitchBox.MinHeight = 120
    [void]$pitchCard.Body.Children.Add($pitchBox)

    $stressCard = New-CardBorder -Title 'Stress test scenarios'
    $stressBox = [System.Windows.Controls.TextBox]::new()
    $stressBox.Background = $script:ui.ActiveBrush
    $stressBox.Foreground = $script:ui.ForegroundBrush
    $stressBox.BorderBrush = $script:ui.BorderBrush
    $stressBox.BorderThickness = [System.Windows.Thickness]::new(1)
    $stressBox.FontFamily = 'Consolas'
    $stressBox.FontSize = 13
    $stressBox.AcceptsReturn = $true
    $stressBox.TextWrapping = 'Wrap'
    $stressBox.IsReadOnly = $true
    $stressBox.MinHeight = 150
    [void]$stressCard.Body.Children.Add($stressBox)

    $detailsCard = New-CardBorder -Title 'Detailed breakdown'
    $detailsBox = [System.Windows.Controls.TextBox]::new()
    $detailsBox.Background = $script:ui.ActiveBrush
    $detailsBox.Foreground = $script:ui.ForegroundBrush
    $detailsBox.BorderBrush = $script:ui.BorderBrush
    $detailsBox.BorderThickness = [System.Windows.Thickness]::new(1)
    $detailsBox.FontFamily = 'Consolas'
    $detailsBox.FontSize = 13
    $detailsBox.AcceptsReturn = $true
    $detailsBox.TextWrapping = 'Wrap'
    $detailsBox.IsReadOnly = $true
    $detailsBox.VerticalScrollBarVisibility = 'Auto'
    $detailsBox.MinHeight = 420
    [void]$detailsCard.Body.Children.Add($detailsBox)

    [void]$ResultPanel.Children.Add($headerCard.Border)
    [void]$ResultPanel.Children.Add($execCard.Border)
    [void]$ResultPanel.Children.Add($pitchCard.Border)
    [void]$ResultPanel.Children.Add($stressCard.Border)
    [void]$ResultPanel.Children.Add($detailsCard.Border)

    return @{ 
        Badge = $badge
        BadgeText = $badgeText
        SummaryGrid = $summaryGrid
        Note = $note
        ExecutiveHeadline = $execHeadline
        ExecutiveBody = $execBody
        PitchBox = $pitchBox
        StressBox = $stressBox
        DetailsBox = $detailsBox
    }
}

function Add-SummaryMetric {
    param(
        [System.Windows.Controls.Panel]$Panel,
        [string]$Label,
        [string]$Value
    )

    $border = [System.Windows.Controls.Border]::new()
    $border.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#172434')
    $border.BorderBrush = $script:ui.BorderBrush
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.Margin = [System.Windows.Thickness]::new(0, 0, 12, 12)
    $border.Padding = [System.Windows.Thickness]::new(14)
    $border.CornerRadius = [System.Windows.CornerRadius]::new(14)

    $stack = [System.Windows.Controls.StackPanel]::new()
    $accent = [System.Windows.Controls.Border]::new()
    $accent.Background = $script:ui.AccentBrush
    $accent.Height = 4
    $accent.CornerRadius = [System.Windows.CornerRadius]::new(4)
    $accent.Margin = [System.Windows.Thickness]::new(0, 0, 0, 10)
    [void]$stack.Children.Add($accent)

    $labelBlock = [System.Windows.Controls.TextBlock]::new()
    $labelBlock.Text = $Label
    $labelBlock.Foreground = $script:ui.MutedBrush
    $labelBlock.TextWrapping = 'Wrap'
    $labelBlock.Margin = [System.Windows.Thickness]::new(0, 0, 0, 4)
    [void]$stack.Children.Add($labelBlock)

    $valueBlock = [System.Windows.Controls.TextBlock]::new()
    $valueBlock.Text = $Value
    $valueBlock.Foreground = $script:ui.ForegroundBrush
    $valueBlock.FontSize = 20
    $valueBlock.FontWeight = 'SemiBold'
    [void]$stack.Children.Add($valueBlock)

    $border.Child = $stack
    [void]$Panel.Children.Add($border)
}

function Set-ResultView {
    param(
        [hashtable]$ResultView,
        [hashtable]$Status,
        [array]$Metrics,
        [string[]]$DetailLines,
        [string]$Note,
        [string]$ExecutiveHeadline,
        [string]$ExecutiveBody,
        [string]$PitchText,
        [string]$StressText
    )

    $ResultView.Badge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Status.Color)
    $ResultView.BadgeText.Text = $Status.Text
    $ResultView.Note.Text = "$($Status.Note) $Note".Trim()
    $ResultView.ExecutiveHeadline.Text = $ExecutiveHeadline
    $ResultView.ExecutiveBody.Text = $ExecutiveBody
    $ResultView.PitchBox.Text = $PitchText
    $ResultView.StressBox.Text = $StressText
    $ResultView.SummaryGrid.Children.Clear()

    foreach ($metric in $Metrics) {
        Add-SummaryMetric -Panel $ResultView.SummaryGrid -Label $metric.Label -Value $metric.Value
    }

    $ResultView.DetailsBox.Text = ($DetailLines -join [Environment]::NewLine)
}

function Reset-ValidationState {
    param([hashtable]$Page)

    foreach ($entry in $Page.Meta.GetEnumerator()) {
        $control = $Page[$entry.Key]
        $meta = $entry.Value
        switch ($meta.Type) {
            'text' {
                $control.BorderBrush = $script:ui.BorderBrush
                $control.Background = $script:ui.ActiveBrush
            }
            'combo' {
                $control.BorderBrush = $script:ui.BorderBrush
                $control.Background = $script:ui.ActiveBrush
            }
        }
    }
}

function Mark-ControlInvalid {
    param([object]$Control)

    $Control.BorderBrush = $script:ui.DangerBrush
    $Control.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#2A1B1B')
}

function Read-NumberField {
    param(
        [hashtable]$Page,
        [string]$Key,
        [switch]$Percent,
        [double]$Min = 0
    )

    $control = $Page[$Key]
    $meta = $Page.Meta[$Key]
    $value = Get-DecimalValue -Text $control.Text

    if ([double]::IsNaN($value) -or $value -lt $Min) {
        Mark-ControlInvalid -Control $control
        $script:validationErrors.Add("$($meta.Label) must be a number greater than or equal to $Min.") | Out-Null
        return $null
    }

    if ($Percent) {
        return $value / 100
    }

    return $value
}

function Reset-PageDefaults {
    param([hashtable]$Page)

    foreach ($entry in $Page.Meta.GetEnumerator()) {
        $control = $Page[$entry.Key]
        $meta = $entry.Value
        switch ($meta.Type) {
            'text' { $control.Text = [string]$meta.Default }
            'combo' { $control.SelectedItem = $meta.Default }
            'check' { $control.IsChecked = [bool]$meta.Default }
        }
    }

    $Page.ShareText = ''
    $Page.ExportRows = @()
    Reset-ValidationState -Page $Page
}

function Show-ValidationResult {
    param(
        [hashtable]$ResultView,
        [string]$Strategy
    )

    Set-ResultView -ResultView $ResultView -Status @{ Text = 'Input error'; Color = '#E57373'; Note = 'Please fix the highlighted fields and run the calculation again.' } -Metrics @(
        @{ Label = 'Strategy'; Value = $Strategy },
        @{ Label = 'Errors'; Value = [string]$script:validationErrors.Count }
    ) -DetailLines $script:validationErrors.ToArray() -Note '' -ExecutiveHeadline 'Input validation required' -ExecutiveBody 'Some required numeric fields are missing or invalid. Correct the highlighted inputs and rerun the deal.' -PitchText 'Validation failed. Fix the highlighted fields before sharing this deal summary.' -StressText 'Stress scenarios are available after a successful calculation.'
    Write-Log -Message ($script:validationErrors -join ' ') -Level 'Warn'
    Set-BusyLabel -Text 'Validation error'
}

function Get-LeadSnapshot {
    return [ordered]@{
        Name = [string]$script:ui.LeadName.Text
        Email = [string]$script:ui.LeadEmail.Text
        Phone = [string]$script:ui.LeadPhone.Text
        Source = [string]$script:ui.LeadSource.Text
    }
}

function Set-LeadSnapshot {
    param([object]$Lead)

    if (-not $Lead) { return }
    $script:ui.LeadName.Text = [string]$Lead.Name
    $script:ui.LeadEmail.Text = [string]$Lead.Email
    $script:ui.LeadPhone.Text = [string]$Lead.Phone
    $script:ui.LeadSource.Text = [string]$Lead.Source
}

function Get-DisplayLeadName {
    $name = [string]$script:ui.LeadName.Text
    if ([string]::IsNullOrWhiteSpace($name) -or $name -eq 'Client / investor name') { return 'Investor' }
    return $name
}

function New-StressLine {
    param(
        [string]$Label,
        [double]$MonthlyCashflow,
        [double]$AnnualReturn,
        [string]$ReturnLabel
    )

    return ('{0,-26} Cashflow {1,12} | {2}: {3}' -f $Label, (Format-Currency $MonthlyCashflow), $ReturnLabel, (Format-Percent $AnnualReturn))
}

function Get-BtlStressText {
    param([hashtable]$Data)

    $base = Get-BtlAnalysis -Data $Data
    $rateUp = $Data.Clone(); $rateUp.InterestRate += 0.01
    $rentDown = $Data.Clone(); $rentDown.MonthlyRent *= 0.90
    $costUp = $Data.Clone(); $costUp.RefurbCost *= 1.10

    $a = Get-BtlAnalysis -Data $rateUp
    $b = Get-BtlAnalysis -Data $rentDown
    $c = Get-BtlAnalysis -Data $costUp

    return @(
        'BTL STRESS TEST'
        ''
        (New-StressLine -Label 'Base case' -MonthlyCashflow $base.MonthlyCashflow -AnnualReturn $base.CashOnCash -ReturnLabel 'CoC')
        (New-StressLine -Label 'Interest rate +1.00%' -MonthlyCashflow $a.MonthlyCashflow -AnnualReturn $a.CashOnCash -ReturnLabel 'CoC')
        (New-StressLine -Label 'Rent -10%' -MonthlyCashflow $b.MonthlyCashflow -AnnualReturn $b.CashOnCash -ReturnLabel 'CoC')
        (New-StressLine -Label 'Refurb +10%' -MonthlyCashflow $c.MonthlyCashflow -AnnualReturn $c.CashOnCash -ReturnLabel 'CoC')
    ) -join [Environment]::NewLine
}

function Get-BrrrStressText {
    param([hashtable]$Data)

    $base = Get-BrrrAnalysis -Data $Data
    $rateUp = $Data.Clone(); $rateUp.RefinanceInterestRate += 0.01
    $rentDown = $Data.Clone(); $rentDown.MonthlyRent *= 0.90
    $valueDown = $Data.Clone(); $valueDown.EndValue *= 0.95

    $a = Get-BrrrAnalysis -Data $rateUp
    $b = Get-BrrrAnalysis -Data $rentDown
    $c = Get-BrrrAnalysis -Data $valueDown

    return @(
        'BRRR STRESS TEST'
        ''
        (New-StressLine -Label 'Base case' -MonthlyCashflow $base.MonthlyCashflow -AnnualReturn $base.ROCE -ReturnLabel 'ROCE')
        (New-StressLine -Label 'Refi rate +1.00%' -MonthlyCashflow $a.MonthlyCashflow -AnnualReturn $a.ROCE -ReturnLabel 'ROCE')
        (New-StressLine -Label 'Rent -10%' -MonthlyCashflow $b.MonthlyCashflow -AnnualReturn $b.ROCE -ReturnLabel 'ROCE')
        (New-StressLine -Label 'End value -5%' -MonthlyCashflow $c.MonthlyCashflow -AnnualReturn $c.ROCE -ReturnLabel 'ROCE')
    ) -join [Environment]::NewLine
}

function Get-HmoStressText {
    param([hashtable]$Data)

    $base = Get-HmoAnalysis -Data $Data
    $rateUp = $Data.Clone(); $rateUp.MortgageRate += 0.01
    $rentDown = $Data.Clone(); $rentDown.RoomRents = @($Data.RoomRents | ForEach-Object { $_ * 0.90 })
    $billsUp = $Data.Clone(); $billsUp.MonthlyUtilities *= 1.15

    $a = Get-HmoAnalysis -Data $rateUp
    $b = Get-HmoAnalysis -Data $rentDown
    $c = Get-HmoAnalysis -Data $billsUp

    return @(
        'HMO STRESS TEST'
        ''
        (New-StressLine -Label 'Base case' -MonthlyCashflow $base.MonthlyCashflow -AnnualReturn $base.ROI -ReturnLabel 'ROI')
        (New-StressLine -Label 'Mortgage rate +1.00%' -MonthlyCashflow $a.MonthlyCashflow -AnnualReturn $a.ROI -ReturnLabel 'ROI')
        (New-StressLine -Label 'Room rents -10%' -MonthlyCashflow $b.MonthlyCashflow -AnnualReturn $b.ROI -ReturnLabel 'ROI')
        (New-StressLine -Label 'Bills +15%' -MonthlyCashflow $c.MonthlyCashflow -AnnualReturn $c.ROI -ReturnLabel 'ROI')
    ) -join [Environment]::NewLine
}

function Get-PageInputSnapshot {
    param([hashtable]$Page)

    $snapshot = [ordered]@{}
    foreach ($entry in ($Page.Meta.GetEnumerator() | Sort-Object Key)) {
        $meta = $entry.Value
        $control = $Page[$entry.Key]
        switch ($meta.Type) {
            'check' { $snapshot[$entry.Key] = [bool]$control.IsChecked }
            'combo' { $snapshot[$entry.Key] = [string]$control.SelectedItem }
            default { $snapshot[$entry.Key] = [string]$control.Text }
        }
    }
    return $snapshot
}

function Save-PageDeal {
    param(
        [hashtable]$Page,
        [string]$Strategy
    )

    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Save deal'
    $dialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
    $dialog.DefaultExt = 'json'
    $dialog.FileName = ('{0}-deal-{1}.json' -f $Strategy.ToLower(), (Get-Date -Format 'yyyyMMdd-HHmmss'))
    if (-not $dialog.ShowDialog()) { return }

    $payload = [ordered]@{
        Strategy = $Strategy
        SavedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Lead = Get-LeadSnapshot
        Inputs = Get-PageInputSnapshot -Page $Page
        ExportRows = @($Page.ExportRows)
        ShareText = $Page.ShareText
    }

    $payload | ConvertTo-Json -Depth 6 | Set-Content -Path $dialog.FileName -Encoding UTF8
    Write-Log -Message ("$Strategy deal saved: {0}" -f $dialog.FileName)
    Set-BusyLabel -Text "$Strategy saved"
}

function Load-PageDeal {
    param(
        [hashtable]$Page,
        [string]$Strategy,
        [scriptblock]$RecalculateAction
    )

    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Title = 'Load deal'
    $dialog.Filter = 'JSON files (*.json)|*.json|All files (*.*)|*.*'
    if (-not $dialog.ShowDialog()) { return }

    $payload = Get-Content -Path $dialog.FileName -Raw | ConvertFrom-Json
    if ($payload.Strategy -ne $Strategy) {
        [System.Windows.MessageBox]::Show(
            "This file contains a $($payload.Strategy) deal, not $Strategy.",
            'UK Property Investment Calculator',
            'OK',
            'Warning'
        ) | Out-Null
        return
    }

    Set-LeadSnapshot -Lead $payload.Lead
    foreach ($entry in $Page.Meta.GetEnumerator()) {
        $key = $entry.Key
        if (-not ($payload.Inputs.PSObject.Properties.Name -contains $key)) { continue }
        $control = $Page[$key]
        $meta = $entry.Value
        $value = $payload.Inputs.$key
        switch ($meta.Type) {
            'check' { $control.IsChecked = [bool]$value }
            'combo' { $control.SelectedItem = [string]$value }
            default { $control.Text = [string]$value }
        }
    }

    & $RecalculateAction
    Write-Log -Message ("$Strategy deal loaded: {0}" -f $dialog.FileName)
    Set-BusyLabel -Text "$Strategy loaded"
}

function Copy-PageSummary {
    param(
        [hashtable]$Page,
        [string]$Strategy
    )

    if ([string]::IsNullOrWhiteSpace($Page.ShareText)) {
        [System.Windows.MessageBox]::Show(
            'Calculate the deal first, then copy the summary.',
            'UK Property Investment Calculator',
            'OK',
            'Information'
        ) | Out-Null
        return
    }

    [System.Windows.Clipboard]::SetText($Page.ShareText)
    Write-Log -Message ("$Strategy summary copied to clipboard.")
    Set-BusyLabel -Text "$Strategy summary copied"
}

function Set-PageExportData {
    param(
        [hashtable]$Page,
        [string]$Strategy,
        [hashtable]$InputData,
        [hashtable]$ResultData,
        [hashtable]$StatusData
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    foreach ($entry in $Page.Meta.GetEnumerator() | Sort-Object Key) {
        $meta = $entry.Value
        if ($meta.Type -eq 'check') {
            $value = [bool]$Page[$entry.Key].IsChecked
        }
        elseif ($meta.Type -eq 'combo') {
            $value = [string]$Page[$entry.Key].SelectedItem
        }
        else {
            $value = [string]$Page[$entry.Key].Text
        }

        $rows.Add([pscustomobject]@{
            ExportedAt = $timestamp
            Strategy = $Strategy
            Section = 'Input'
            Field = $meta.Label
            Key = $entry.Key
            Value = $value
        }) | Out-Null
    }

    foreach ($entry in $ResultData.GetEnumerator()) {
        $rows.Add([pscustomobject]@{
            ExportedAt = $timestamp
            Strategy = $Strategy
            Section = 'Result'
            Field = $entry.Key
            Key = $entry.Key
            Value = [string]$entry.Value
        }) | Out-Null
    }

    foreach ($entry in $StatusData.GetEnumerator()) {
        $rows.Add([pscustomobject]@{
            ExportedAt = $timestamp
            Strategy = $Strategy
            Section = 'Status'
            Field = $entry.Key
            Key = $entry.Key
            Value = [string]$entry.Value
        }) | Out-Null
    }

    $Page.ExportRows = $rows
}

function Export-PageCsv {
    param(
        [hashtable]$Page,
        [string]$Strategy
    )

    if (-not $Page.ExportRows -or $Page.ExportRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'Calculate the deal first, then export it to CSV.',
            'UK Property Investment Calculator',
            'OK',
            'Information'
        ) | Out-Null
        return
    }

    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Export deal to CSV'
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dialog.DefaultExt = 'csv'
    $dialog.FileName = ('{0}-{1}.csv' -f $Strategy, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    if (-not $dialog.ShowDialog()) {
        return
    }

    $Page.ExportRows | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
    Write-Log -Message ("$Strategy exported to CSV: {0}" -f $dialog.FileName)
    Set-BusyLabel -Text "$Strategy exported"
}

function Export-PageXlsx {
    param(
        [hashtable]$Page,
        [string]$Strategy
    )

    if (-not $Page.ExportRows -or $Page.ExportRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'Calculate the deal first, then export it to XLSX.',
            'UK Property Investment Calculator',
            'OK',
            'Information'
        ) | Out-Null
        return
    }

    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Export deal to XLSX'
    $dialog.Filter = 'Excel workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*'
    $dialog.DefaultExt = 'xlsx'
    $dialog.FileName = ('{0}-{1}.xlsx' -f $Strategy, (Get-Date -Format 'yyyyMMdd-HHmmss'))

    if (-not $dialog.ShowDialog()) {
        return
    }

    $excel = $null
    $workbook = $null
    $worksheet = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = 'Deal Export'

        $headers = @('ExportedAt', 'Strategy', 'Section', 'Field', 'Key', 'Value')
        for ($i = 0; $i -lt $headers.Count; $i++) {
            $worksheet.Cells.Item(1, $i + 1) = $headers[$i]
        }

        $rowIndex = 2
        foreach ($row in $Page.ExportRows) {
            $worksheet.Cells.Item($rowIndex, 1) = $row.ExportedAt
            $worksheet.Cells.Item($rowIndex, 2) = $row.Strategy
            $worksheet.Cells.Item($rowIndex, 3) = $row.Section
            $worksheet.Cells.Item($rowIndex, 4) = $row.Field
            $worksheet.Cells.Item($rowIndex, 5) = $row.Key
            $worksheet.Cells.Item($rowIndex, 6) = $row.Value
            $rowIndex++
        }

        $headerRange = $worksheet.Range('A1', 'F1')
        $headerRange.Font.Bold = $true
        $headerRange.Interior.Color = 39423
        $headerRange.Font.Color = 16777215
        $worksheet.Columns.AutoFit() | Out-Null
        $worksheet.Application.ActiveWindow.SplitRow = 1
        $worksheet.Application.ActiveWindow.FreezePanes = $true

        $workbook.SaveAs($dialog.FileName, 51)
        Write-Log -Message ("$Strategy exported to XLSX: {0}" -f $dialog.FileName)
        Set-BusyLabel -Text "$Strategy exported"
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to export XLSX.`n`n$($_.Exception.Message)",
            'UK Property Investment Calculator',
            'OK',
            'Error'
        ) | Out-Null
        throw
    }
    finally {
        if ($workbook) { $workbook.Close($false) }
        if ($excel) { $excel.Quit() }
        foreach ($comObject in @($headerRange, $worksheet, $workbook, $excel)) {
            if ($null -ne $comObject) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Export-PagePdf {
    param(
        [hashtable]$Page,
        [string]$Strategy
    )

    if (-not $Page.ExportRows -or $Page.ExportRows.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            'Calculate the deal first, then export it to PDF.',
            'UK Property Investment Calculator',
            'OK',
            'Information'
        ) | Out-Null
        return
    }

    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Title = 'Export investor PDF report'
    $dialog.Filter = 'PDF files (*.pdf)|*.pdf|All files (*.*)|*.*'
    $dialog.DefaultExt = 'pdf'
    $dialog.FileName = ('{0}-investor-report-{1}.pdf' -f $Strategy.ToLower(), (Get-Date -Format 'yyyyMMdd-HHmmss'))
    if (-not $dialog.ShowDialog()) { return }

    $lead = Get-LeadSnapshot
    $word = $null
    $document = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.ScreenUpdating = $false
        $document = $word.Documents.Add()
        $selection = $word.Selection

        $selection.Style = 'Title'
        $selection.TypeText('UK Property Investment Calculator')
        $selection.TypeParagraph()
        $selection.Style = 'Subtitle'
        $selection.TypeText("$Strategy Investor Report")
        $selection.TypeParagraph()
        $selection.TypeText('Prepared by marcinsakowski.com')
        $selection.TypeParagraph()
        $selection.TypeText(('Prepared for: {0}' -f (Get-DisplayLeadName)))
        $selection.TypeParagraph()
        $selection.TypeText(('Generated: {0}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
        $selection.TypeParagraph()
        $selection.InsertParagraphAfter()

        $selection.Style = 'Heading 1'
        $selection.TypeText('Executive Summary')
        $selection.TypeParagraph()
        $selection.Style = 'Normal'
        $selection.TypeText($Page.ResultView.ExecutiveHeadline.Text)
        $selection.TypeParagraph()
        $selection.TypeText($Page.ResultView.ExecutiveBody.Text)
        $selection.TypeParagraph()
        $selection.TypeText($Page.ResultView.Note.Text)
        $selection.TypeParagraph()
        $selection.InsertParagraphAfter()

        $selection.Style = 'Heading 1'
        $selection.TypeText('Lead Details')
        $selection.TypeParagraph()
        $selection.Style = 'Normal'
        foreach ($line in @(
            ('Name: ' + $lead.Name),
            ('Email: ' + $lead.Email),
            ('Phone: ' + $lead.Phone),
            ('Source / Notes: ' + $lead.Source)
        )) {
            $selection.TypeText($line)
            $selection.TypeParagraph()
        }
        $selection.InsertParagraphAfter()

        $selection.Style = 'Heading 1'
        $selection.TypeText('Share-Ready Pitch')
        $selection.TypeParagraph()
        $selection.Style = 'Normal'
        $selection.TypeText($Page.ShareText)
        $selection.TypeParagraph()
        $selection.InsertParagraphAfter()

        $selection.Style = 'Heading 1'
        $selection.TypeText('Stress Test Scenarios')
        $selection.TypeParagraph()
        $selection.Style = 'Normal'
        $selection.TypeText($Page.ResultView.StressBox.Text)
        $selection.TypeParagraph()
        $selection.InsertParagraphAfter()

        $selection.Style = 'Heading 1'
        $selection.TypeText('Detailed Breakdown')
        $selection.TypeParagraph()
        $selection.Style = 'Normal'
        $selection.TypeText($Page.ResultView.DetailsBox.Text)
        $selection.TypeParagraph()
        $selection.InsertParagraphAfter()

        $selection.Style = 'Heading 1'
        $selection.TypeText('Inputs and Outputs')
        $selection.TypeParagraph()
        $selection.Style = 'Normal'
        foreach ($row in $Page.ExportRows) {
            $selection.TypeText(('{0} | {1} | {2}: {3}' -f $row.Section, $row.Field, $row.Key, $row.Value))
            $selection.TypeParagraph()
        }

        $document.ExportAsFixedFormat($dialog.FileName, 17)
        Write-Log -Message ("$Strategy exported to PDF: {0}" -f $dialog.FileName)
        Set-BusyLabel -Text "$Strategy exported"
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to export PDF.`n`n$($_.Exception.Message)",
            'UK Property Investment Calculator',
            'OK',
            'Error'
        ) | Out-Null
        throw
    }
    finally {
        if ($document) { $document.Close($false) }
        if ($word) { $word.Quit() }
        foreach ($comObject in @($selection, $document, $word)) {
            if ($null -ne $comObject) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject)
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Build-Page {
    param(
        [string]$Name,
        [System.Windows.Controls.Panel]$FormPanel,
        [array]$Cards,
        [System.Windows.Controls.Panel]$ResultPanel,
        [string]$ResultDescription
    )

    $page = @{ Meta = @{} }
    foreach ($card in $Cards) {
        $cardBlock = New-CardBorder -Title $card.Title
        Add-FieldsGrid -Parent $cardBlock.Body -Fields $card.Fields -Store $page -Window $script:ui.Window
        [void]$FormPanel.Children.Add($cardBlock.Border)
    }

    $page.ResultView = Initialize-ResultPanel -ResultPanel $ResultPanel -Strategy $Name -Description $ResultDescription
    $page.ShareText = ''
    $script:pages[$Name] = $page
}

$regionOptions = @('England / NI', 'Wales', 'Manual (incl. Scotland)')

Build-Page -Name 'BTL' -FormPanel $script:ui.BTLFormHost -ResultPanel $script:ui.BTLResultHost -ResultDescription 'Screen a traditional buy-to-let using purchase cost, finance, rent, voids and operating cost assumptions.' -Cards @(
    @{
        Title = 'Deal setup'
        Fields = @(
            @{ Key = 'Region'; Label = 'Tax region'; Type = 'combo'; Options = $regionOptions; Default = 'England / NI' },
            @{ Key = 'IsAdditional'; Label = 'Additional property / investor purchase'; Type = 'check'; Default = $true },
            @{ Key = 'ManualStampDuty'; Label = 'Manual stamp duty'; Type = 'text'; Default = '0' },
            @{ Key = 'PurchasePrice'; Label = 'Purchase price'; Type = 'text'; Default = '85000' },
            @{ Key = 'DepositPct'; Label = 'Deposit %'; Type = 'text'; Default = '25' },
            @{ Key = 'InterestRate'; Label = 'Interest rate %'; Type = 'text'; Default = '5.5' },
            @{ Key = 'FinanceMonths'; Label = 'Finance months before let'; Type = 'text'; Default = '3' }
        )
    },
    @{
        Title = 'Costs and rent'
        Fields = @(
            @{ Key = 'LegalFees'; Label = 'Legal fees'; Type = 'text'; Default = '1000' },
            @{ Key = 'BrokerFees'; Label = 'Broker / lender fees'; Type = 'text'; Default = '500' },
            @{ Key = 'SurveyFees'; Label = 'Survey fees'; Type = 'text'; Default = '300' },
            @{ Key = 'SourcingFees'; Label = 'Sourcing / agent fee'; Type = 'text'; Default = '0' },
            @{ Key = 'OtherBuyingCosts'; Label = 'Other buying costs'; Type = 'text'; Default = '0' },
            @{ Key = 'RefurbCost'; Label = 'Refurb cost'; Type = 'text'; Default = '5000' },
            @{ Key = 'VacantMonthlyCosts'; Label = 'Vacant monthly holding costs'; Type = 'text'; Default = '100' },
            @{ Key = 'VacantMonths'; Label = 'Vacant months'; Type = 'text'; Default = '3' },
            @{ Key = 'MonthlyOperatingCosts'; Label = 'Monthly operating costs when let'; Type = 'text'; Default = '95' },
            @{ Key = 'MonthlyRent'; Label = 'Monthly rent'; Type = 'text'; Default = '750' }
        )
    }
)

Build-Page -Name 'BRRR' -FormPanel $script:ui.BRRRFormHost -ResultPanel $script:ui.BRRRResultHost -ResultDescription 'Analyse a BRRR deal across purchase, refurb, bridge period, refinance and post-refinance performance.' -Cards @(
    @{
        Title = 'Acquisition and refinance'
        Fields = @(
            @{ Key = 'Region'; Label = 'Tax region'; Type = 'combo'; Options = $regionOptions; Default = 'England / NI' },
            @{ Key = 'IsAdditional'; Label = 'Additional property / investor purchase'; Type = 'check'; Default = $true },
            @{ Key = 'ManualStampDuty'; Label = 'Manual stamp duty'; Type = 'text'; Default = '0' },
            @{ Key = 'PurchasePrice'; Label = 'Purchase price'; Type = 'text'; Default = '97000' },
            @{ Key = 'PurchaseDepositPct'; Label = 'Purchase deposit %'; Type = 'text'; Default = '100' },
            @{ Key = 'PurchaseInterestRate'; Label = 'Purchase / bridge rate %'; Type = 'text'; Default = '8' },
            @{ Key = 'FinanceMonths'; Label = 'Finance months to refinance'; Type = 'text'; Default = '6' },
            @{ Key = 'EndValue'; Label = 'End value / revaluation'; Type = 'text'; Default = '130000' },
            @{ Key = 'RefinanceDepositPct'; Label = 'Refinance deposit %'; Type = 'text'; Default = '25' },
            @{ Key = 'RefinanceInterestRate'; Label = 'Refinance interest rate %'; Type = 'text'; Default = '5.25' },
            @{ Key = 'RefinanceCosts'; Label = 'Refinance costs'; Type = 'text'; Default = '800' }
        )
    },
    @{
        Title = 'Costs and rent'
        Fields = @(
            @{ Key = 'LegalFees'; Label = 'Legal fees'; Type = 'text'; Default = '1000' },
            @{ Key = 'BrokerFees'; Label = 'Broker / lender fees'; Type = 'text'; Default = '500' },
            @{ Key = 'SurveyFees'; Label = 'Survey fees'; Type = 'text'; Default = '300' },
            @{ Key = 'SourcingFees'; Label = 'Sourcing / agent fee'; Type = 'text'; Default = '0' },
            @{ Key = 'OtherBuyingCosts'; Label = 'Other buying costs'; Type = 'text'; Default = '0' },
            @{ Key = 'RefurbCost'; Label = 'Refurb cost'; Type = 'text'; Default = '10000' },
            @{ Key = 'VacantMonthlyCosts'; Label = 'Vacant monthly holding costs'; Type = 'text'; Default = '100' },
            @{ Key = 'VacantMonths'; Label = 'Vacant months'; Type = 'text'; Default = '3' },
            @{ Key = 'MonthlyOperatingCosts'; Label = 'Monthly operating costs when let'; Type = 'text'; Default = '95' },
            @{ Key = 'MonthlyRent'; Label = 'Monthly rent'; Type = 'text'; Default = '695' }
        )
    }
)

Build-Page -Name 'HMO' -FormPanel $script:ui.HMOFormHost -ResultPanel $script:ui.HMOResultHost -ResultDescription 'Check a rent-by-room HMO using rent roll, bills, management, refinance and commercial yield.' -Cards @(
    @{
        Title = 'Deal setup'
        Fields = @(
            @{ Key = 'Region'; Label = 'Tax region'; Type = 'combo'; Options = $regionOptions; Default = 'England / NI' },
            @{ Key = 'IsAdditional'; Label = 'Additional property / investor purchase'; Type = 'check'; Default = $true },
            @{ Key = 'ManualStampDuty'; Label = 'Manual stamp duty'; Type = 'text'; Default = '0' },
            @{ Key = 'PurchasePrice'; Label = 'Purchase price'; Type = 'text'; Default = '180000' },
            @{ Key = 'SourcingFee'; Label = 'Sourcing fee'; Type = 'text'; Default = '3000' },
            @{ Key = 'DevelopmentCost'; Label = 'Development / refurb'; Type = 'text'; Default = '40000' },
            @{ Key = 'ProjectManagement'; Label = 'Project management'; Type = 'text'; Default = '5000' },
            @{ Key = 'FurniturePurchased'; Label = 'Furniture purchased'; Type = 'text'; Default = '6000' },
            @{ Key = 'PurchaseCosts'; Label = 'Other purchase costs'; Type = 'text'; Default = '2500' }
        )
    },
    @{
        Title = 'Refinance and operating model'
        Fields = @(
            @{ Key = 'RefinanceValue'; Label = 'Revalued price'; Type = 'text'; Default = '290000' },
            @{ Key = 'LtvPct'; Label = 'LTV %'; Type = 'text'; Default = '75' },
            @{ Key = 'MortgageRate'; Label = 'Mortgage rate %'; Type = 'text'; Default = '5.5' },
            @{ Key = 'ManagementFeePct'; Label = 'Management fee %'; Type = 'text'; Default = '12' },
            @{ Key = 'YieldTarget'; Label = 'Commercial yield target %'; Type = 'text'; Default = '12' },
            @{ Key = 'MonthlyOperatingCosts'; Label = 'Monthly operating costs'; Type = 'text'; Default = '250' },
            @{ Key = 'MonthlyUtilities'; Label = 'Monthly utilities / bills'; Type = 'text'; Default = '650' },
            @{ Key = 'MonthlyFurnitureLease'; Label = 'Monthly furniture lease'; Type = 'text'; Default = '0' },
            @{ Key = 'InvestorFunds'; Label = 'Investor funds'; Type = 'text'; Default = '0' },
            @{ Key = 'InvestorRate'; Label = 'Investor rate %'; Type = 'text'; Default = '8' }
        )
    },
    @{
        Title = 'Room rents'
        Fields = @(
            @{ Key = 'Room1'; Label = 'Room 1 rent'; Type = 'text'; Default = '650' },
            @{ Key = 'Room2'; Label = 'Room 2 rent'; Type = 'text'; Default = '650' },
            @{ Key = 'Room3'; Label = 'Room 3 rent'; Type = 'text'; Default = '650' },
            @{ Key = 'Room4'; Label = 'Room 4 rent'; Type = 'text'; Default = '650' },
            @{ Key = 'Room5'; Label = 'Room 5 rent'; Type = 'text'; Default = '600' },
            @{ Key = 'Room6'; Label = 'Room 6 rent'; Type = 'text'; Default = '600' }
        )
    }
)

function Show-Page {
    param([string]$Name)

    $script:ui.BTLPage.Visibility = 'Collapsed'
    $script:ui.BRRRPage.Visibility = 'Collapsed'
    $script:ui.HMOPage.Visibility = 'Collapsed'

    $script:ui.NavBTL.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1E2836')
    $script:ui.NavBRRR.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1E2836')
    $script:ui.NavHMO.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1E2836')
    $script:ui.NavBTL.Foreground = $script:ui.ForegroundBrush
    $script:ui.NavBRRR.Foreground = $script:ui.ForegroundBrush
    $script:ui.NavHMO.Foreground = $script:ui.ForegroundBrush

    switch ($Name) {
        'BTL' {
            $script:ui.BTLPage.Visibility = 'Visible'
            $script:ui.NavBTL.Background = $script:ui.HeroBrush
            $script:ui.NavBTL.Foreground = [System.Windows.Media.Brushes]::White
            $script:ui.HeaderSubtitle.Text = 'Traditional buy-to-let deal analysis focused on yield and cashflow.'
        }
        'BRRR' {
            $script:ui.BRRRPage.Visibility = 'Visible'
            $script:ui.NavBRRR.Background = $script:ui.HeroBrush
            $script:ui.NavBRRR.Foreground = [System.Windows.Media.Brushes]::White
            $script:ui.HeaderSubtitle.Text = 'BRRR analysis with purchase funding, refurb, refinance and cash left in the deal.'
        }
        'HMO' {
            $script:ui.HMOPage.Visibility = 'Visible'
            $script:ui.NavHMO.Background = $script:ui.HeroBrush
            $script:ui.NavHMO.Foreground = [System.Windows.Media.Brushes]::White
            $script:ui.HeaderSubtitle.Text = 'HMO analysis with room rents, bills, refinance and commercial valuation estimate.'
        }
    }

    Set-BusyLabel -Text "$Name page ready"
}

function Invoke-BTLCalculation {
    $page = $script:pages.BTL
    Reset-ValidationState -Page $page
    $script:validationErrors = [System.Collections.Generic.List[string]]::new()

    $data = @{
        Region = [string]$page.Region.SelectedItem
        IsAdditional = [bool]$page.IsAdditional.IsChecked
        ManualStampDuty = Read-NumberField -Page $page -Key 'ManualStampDuty'
        PurchasePrice = Read-NumberField -Page $page -Key 'PurchasePrice'
        DepositPct = Read-NumberField -Page $page -Key 'DepositPct' -Percent
        InterestRate = Read-NumberField -Page $page -Key 'InterestRate' -Percent
        FinanceMonths = Read-NumberField -Page $page -Key 'FinanceMonths'
        LegalFees = Read-NumberField -Page $page -Key 'LegalFees'
        BrokerFees = Read-NumberField -Page $page -Key 'BrokerFees'
        SurveyFees = Read-NumberField -Page $page -Key 'SurveyFees'
        SourcingFees = Read-NumberField -Page $page -Key 'SourcingFees'
        OtherBuyingCosts = Read-NumberField -Page $page -Key 'OtherBuyingCosts'
        RefurbCost = Read-NumberField -Page $page -Key 'RefurbCost'
        VacantMonthlyCosts = Read-NumberField -Page $page -Key 'VacantMonthlyCosts'
        VacantMonths = Read-NumberField -Page $page -Key 'VacantMonths'
        MonthlyOperatingCosts = Read-NumberField -Page $page -Key 'MonthlyOperatingCosts'
        MonthlyRent = Read-NumberField -Page $page -Key 'MonthlyRent'
    }

    if ($script:validationErrors.Count -gt 0) {
        Show-ValidationResult -ResultView $page.ResultView -Strategy 'BTL'
        return
    }

    $result = Get-BtlAnalysis -Data $data
    $status = Get-DealStatus -Strategy 'BTL' -Result $result

    Set-ResultView -ResultView $page.ResultView -Status $status -Metrics @(
        @{ Label = 'Monthly cashflow'; Value = Format-Currency $result.MonthlyCashflow },
        @{ Label = 'Annual cashflow'; Value = Format-Currency $result.AnnualCashflow },
        @{ Label = 'Gross yield'; Value = Format-Percent $result.GrossYield },
        @{ Label = 'Cash-on-cash return'; Value = Format-Percent $result.CashOnCash },
        @{ Label = 'Total cash in'; Value = Format-Currency $result.TotalCashIn },
        @{ Label = 'Break-even occupancy'; Value = Format-Percent $result.BreakEvenOccupancy }
    ) -DetailLines @(
        'BTL BREAKDOWN'
        ''
        ('Stamp duty:            ' + (Format-Currency $result.StampDuty))
        ('Buying costs:          ' + (Format-Currency $result.BuyingCosts))
        ('Deposit amount:        ' + (Format-Currency $result.DepositAmount))
        ('Loan amount:           ' + (Format-Currency $result.LoanAmount))
        ('Monthly mortgage:      ' + (Format-Currency $result.MonthlyMortgage))
        ('Finance cost:          ' + (Format-Currency $result.FinanceCost))
        ('Holding cost:          ' + (Format-Currency $result.HoldingCost))
        ('Upfront cash:          ' + (Format-Currency $result.UpfrontCash))
        ('Total cash in:         ' + (Format-Currency $result.TotalCashIn))
        ''
        ('Monthly cashflow:      ' + (Format-Currency $result.MonthlyCashflow))
        ('Annual cashflow:       ' + (Format-Currency $result.AnnualCashflow))
        ('Gross yield:           ' + (Format-Percent $result.GrossYield))
        ('Cash-on-cash return:   ' + (Format-Percent $result.CashOnCash))
        ('Break-even occupancy:  ' + (Format-Percent $result.BreakEvenOccupancy))
    ) -Note 'Use manual tax mode if your transaction falls outside England / NI or Wales standard rules.' -ExecutiveHeadline "$($status.Text): BTL cashflow $(Format-Currency $result.MonthlyCashflow) per month" -ExecutiveBody "Prepared for $(Get-DisplayLeadName). This BTL screens at $(Format-Percent $result.GrossYield) gross yield with $(Format-Percent $result.CashOnCash) cash-on-cash return and $(Format-Currency $result.TotalCashIn) total cash required." -PitchText "BTL deal summary`r`nPrepared for: $(Get-DisplayLeadName)`r`nStatus: $($status.Text)`r`nMonthly cashflow: $(Format-Currency $result.MonthlyCashflow)`r`nAnnual cashflow: $(Format-Currency $result.AnnualCashflow)`r`nGross yield: $(Format-Percent $result.GrossYield)`r`nCash-on-cash return: $(Format-Percent $result.CashOnCash)`r`nTotal cash in: $(Format-Currency $result.TotalCashIn)`r`nSource: https://marcinsakowski.com" -StressText (Get-BtlStressText -Data $data)

    Set-PageExportData -Page $page -Strategy 'BTL' -InputData $data -ResultData @{
        StampDuty = Format-Currency $result.StampDuty
        BuyingCosts = Format-Currency $result.BuyingCosts
        DepositAmount = Format-Currency $result.DepositAmount
        LoanAmount = Format-Currency $result.LoanAmount
        MonthlyMortgage = Format-Currency $result.MonthlyMortgage
        FinanceCost = Format-Currency $result.FinanceCost
        HoldingCost = Format-Currency $result.HoldingCost
        UpfrontCash = Format-Currency $result.UpfrontCash
        TotalCashIn = Format-Currency $result.TotalCashIn
        MonthlyCashflow = Format-Currency $result.MonthlyCashflow
        AnnualCashflow = Format-Currency $result.AnnualCashflow
        GrossYield = Format-Percent $result.GrossYield
        CashOnCash = Format-Percent $result.CashOnCash
        BreakEvenOccupancy = Format-Percent $result.BreakEvenOccupancy
    } -StatusData $status
    $page.ShareText = "BTL deal summary`r`nPrepared for: $(Get-DisplayLeadName)`r`nStatus: $($status.Text)`r`nMonthly cashflow: $(Format-Currency $result.MonthlyCashflow)`r`nAnnual cashflow: $(Format-Currency $result.AnnualCashflow)`r`nGross yield: $(Format-Percent $result.GrossYield)`r`nCash-on-cash return: $(Format-Percent $result.CashOnCash)`r`nTotal cash in: $(Format-Currency $result.TotalCashIn)`r`nSource: https://marcinsakowski.com"

    Write-Log -Message ("BTL calculated: cashflow {0}, gross yield {1}." -f (Format-Currency $result.MonthlyCashflow), (Format-Percent $result.GrossYield))
    Set-BusyLabel -Text 'BTL calculated'
}

function Invoke-BRRRCalculation {
    $page = $script:pages.BRRR
    Reset-ValidationState -Page $page
    $script:validationErrors = [System.Collections.Generic.List[string]]::new()

    $data = @{
        Region = [string]$page.Region.SelectedItem
        IsAdditional = [bool]$page.IsAdditional.IsChecked
        ManualStampDuty = Read-NumberField -Page $page -Key 'ManualStampDuty'
        PurchasePrice = Read-NumberField -Page $page -Key 'PurchasePrice'
        PurchaseDepositPct = Read-NumberField -Page $page -Key 'PurchaseDepositPct' -Percent
        PurchaseInterestRate = Read-NumberField -Page $page -Key 'PurchaseInterestRate' -Percent
        FinanceMonths = Read-NumberField -Page $page -Key 'FinanceMonths'
        LegalFees = Read-NumberField -Page $page -Key 'LegalFees'
        BrokerFees = Read-NumberField -Page $page -Key 'BrokerFees'
        SurveyFees = Read-NumberField -Page $page -Key 'SurveyFees'
        SourcingFees = Read-NumberField -Page $page -Key 'SourcingFees'
        OtherBuyingCosts = Read-NumberField -Page $page -Key 'OtherBuyingCosts'
        RefurbCost = Read-NumberField -Page $page -Key 'RefurbCost'
        VacantMonthlyCosts = Read-NumberField -Page $page -Key 'VacantMonthlyCosts'
        VacantMonths = Read-NumberField -Page $page -Key 'VacantMonths'
        MonthlyOperatingCosts = Read-NumberField -Page $page -Key 'MonthlyOperatingCosts'
        MonthlyRent = Read-NumberField -Page $page -Key 'MonthlyRent'
        EndValue = Read-NumberField -Page $page -Key 'EndValue'
        RefinanceDepositPct = Read-NumberField -Page $page -Key 'RefinanceDepositPct' -Percent
        RefinanceInterestRate = Read-NumberField -Page $page -Key 'RefinanceInterestRate' -Percent
        RefinanceCosts = Read-NumberField -Page $page -Key 'RefinanceCosts'
    }

    if ($script:validationErrors.Count -gt 0) {
        Show-ValidationResult -ResultView $page.ResultView -Strategy 'BRRR'
        return
    }

    $result = Get-BrrrAnalysis -Data $data
    $status = Get-DealStatus -Strategy 'BRRR' -Result $result

    Set-ResultView -ResultView $page.ResultView -Status $status -Metrics @(
        @{ Label = 'Money taken out'; Value = Format-Currency $result.MoneyTakenOut },
        @{ Label = 'Cash left in deal'; Value = Format-Currency $result.CashLeftIn },
        @{ Label = 'Monthly cashflow'; Value = Format-Currency $result.MonthlyCashflow },
        @{ Label = 'ROCE'; Value = Format-Percent $result.ROCE },
        @{ Label = 'Yield on purchase'; Value = Format-Percent $result.GrossYieldOnPurchase },
        @{ Label = 'Yield on end value'; Value = Format-Percent $result.GrossYieldOnEndValue }
    ) -DetailLines @(
        'BRRR BREAKDOWN'
        ''
        ('Stamp duty:             ' + (Format-Currency $result.StampDuty))
        ('Buying costs:           ' + (Format-Currency $result.BuyingCosts))
        ('Purchase deposit:       ' + (Format-Currency $result.PurchaseDeposit))
        ('Purchase loan:          ' + (Format-Currency $result.PurchaseLoan))
        ('Monthly bridge cost:    ' + (Format-Currency $result.MonthlyBridgeCost))
        ('Bridge cost total:      ' + (Format-Currency $result.BridgeCostTotal))
        ('Vacant holding total:   ' + (Format-Currency $result.VacantHoldingTotal))
        ('Cash before refinance:  ' + (Format-Currency $result.CashBeforeRefi))
        ('Refinance mortgage:     ' + (Format-Currency $result.RefinanceMortgage))
        ('Money taken out:        ' + (Format-Currency $result.MoneyTakenOut))
        ('Cash left in deal:      ' + (Format-Currency $result.CashLeftIn))
        ('Monthly refi mortgage:  ' + (Format-Currency $result.MonthlyRefiMortgage))
        ('Monthly cashflow:       ' + (Format-Currency $result.MonthlyCashflow))
        ('Annual cashflow:        ' + (Format-Currency $result.AnnualCashflow))
        ('Yield on purchase:      ' + (Format-Percent $result.GrossYieldOnPurchase))
        ('Yield on end value:     ' + (Format-Percent $result.GrossYieldOnEndValue))
        ('ROCE:                   ' + (Format-Percent $result.ROCE))
        ('Equity created:         ' + (Format-Currency $result.EquityCreated))
    ) -Note 'Stress-test the refinance value and lender criteria before relying on the BRRR outcome.' -ExecutiveHeadline "$($status.Text): BRRR leaves $(Format-Currency $result.CashLeftIn) in the deal" -ExecutiveBody "Prepared for $(Get-DisplayLeadName). This BRRR projects $(Format-Currency $result.MoneyTakenOut) returned on refinance, $(Format-Percent $result.ROCE) ROCE, and $(Format-Currency $result.EquityCreated) equity created." -PitchText "BRRR deal summary`r`nPrepared for: $(Get-DisplayLeadName)`r`nStatus: $($status.Text)`r`nMoney taken out: $(Format-Currency $result.MoneyTakenOut)`r`nCash left in: $(Format-Currency $result.CashLeftIn)`r`nMonthly cashflow: $(Format-Currency $result.MonthlyCashflow)`r`nROCE: $(Format-Percent $result.ROCE)`r`nEquity created: $(Format-Currency $result.EquityCreated)`r`nSource: https://marcinsakowski.com" -StressText (Get-BrrrStressText -Data $data)

    Set-PageExportData -Page $page -Strategy 'BRRR' -InputData $data -ResultData @{
        StampDuty = Format-Currency $result.StampDuty
        BuyingCosts = Format-Currency $result.BuyingCosts
        PurchaseDeposit = Format-Currency $result.PurchaseDeposit
        PurchaseLoan = Format-Currency $result.PurchaseLoan
        MonthlyBridgeCost = Format-Currency $result.MonthlyBridgeCost
        BridgeCostTotal = Format-Currency $result.BridgeCostTotal
        VacantHoldingTotal = Format-Currency $result.VacantHoldingTotal
        CashBeforeRefi = Format-Currency $result.CashBeforeRefi
        RefinanceMortgage = Format-Currency $result.RefinanceMortgage
        MoneyTakenOut = Format-Currency $result.MoneyTakenOut
        CashLeftIn = Format-Currency $result.CashLeftIn
        MonthlyRefiMortgage = Format-Currency $result.MonthlyRefiMortgage
        MonthlyCashflow = Format-Currency $result.MonthlyCashflow
        AnnualCashflow = Format-Currency $result.AnnualCashflow
        GrossYieldOnPurchase = Format-Percent $result.GrossYieldOnPurchase
        GrossYieldOnEndValue = Format-Percent $result.GrossYieldOnEndValue
        ROCE = Format-Percent $result.ROCE
        EquityCreated = Format-Currency $result.EquityCreated
    } -StatusData $status
    $page.ShareText = "BRRR deal summary`r`nPrepared for: $(Get-DisplayLeadName)`r`nStatus: $($status.Text)`r`nMoney taken out: $(Format-Currency $result.MoneyTakenOut)`r`nCash left in: $(Format-Currency $result.CashLeftIn)`r`nMonthly cashflow: $(Format-Currency $result.MonthlyCashflow)`r`nROCE: $(Format-Percent $result.ROCE)`r`nEquity created: $(Format-Currency $result.EquityCreated)`r`nSource: https://marcinsakowski.com"

    Write-Log -Message ("BRRR calculated: cash left in {0}, ROCE {1}." -f (Format-Currency $result.CashLeftIn), (Format-Percent $result.ROCE))
    Set-BusyLabel -Text 'BRRR calculated'
}

function Invoke-HMOCalculation {
    $page = $script:pages.HMO
    Reset-ValidationState -Page $page
    $script:validationErrors = [System.Collections.Generic.List[string]]::new()

    $data = @{
        Region = [string]$page.Region.SelectedItem
        IsAdditional = [bool]$page.IsAdditional.IsChecked
        ManualStampDuty = Read-NumberField -Page $page -Key 'ManualStampDuty'
        PurchasePrice = Read-NumberField -Page $page -Key 'PurchasePrice'
        SourcingFee = Read-NumberField -Page $page -Key 'SourcingFee'
        DevelopmentCost = Read-NumberField -Page $page -Key 'DevelopmentCost'
        ProjectManagement = Read-NumberField -Page $page -Key 'ProjectManagement'
        FurniturePurchased = Read-NumberField -Page $page -Key 'FurniturePurchased'
        PurchaseCosts = Read-NumberField -Page $page -Key 'PurchaseCosts'
        RefinanceValue = Read-NumberField -Page $page -Key 'RefinanceValue'
        LtvPct = Read-NumberField -Page $page -Key 'LtvPct' -Percent
        MortgageRate = Read-NumberField -Page $page -Key 'MortgageRate' -Percent
        ManagementFeePct = Read-NumberField -Page $page -Key 'ManagementFeePct' -Percent
        YieldTarget = Read-NumberField -Page $page -Key 'YieldTarget' -Percent -Min 0.01
        MonthlyOperatingCosts = Read-NumberField -Page $page -Key 'MonthlyOperatingCosts'
        MonthlyUtilities = Read-NumberField -Page $page -Key 'MonthlyUtilities'
        MonthlyFurnitureLease = Read-NumberField -Page $page -Key 'MonthlyFurnitureLease'
        InvestorFunds = Read-NumberField -Page $page -Key 'InvestorFunds'
        InvestorRate = Read-NumberField -Page $page -Key 'InvestorRate' -Percent
        RoomRents = @(
            (Read-NumberField -Page $page -Key 'Room1')
            (Read-NumberField -Page $page -Key 'Room2')
            (Read-NumberField -Page $page -Key 'Room3')
            (Read-NumberField -Page $page -Key 'Room4')
            (Read-NumberField -Page $page -Key 'Room5')
            (Read-NumberField -Page $page -Key 'Room6')
        )
    }

    if ($script:validationErrors.Count -gt 0) {
        Show-ValidationResult -ResultView $page.ResultView -Strategy 'HMO'
        return
    }

    $result = Get-HmoAnalysis -Data $data
    $status = Get-DealStatus -Strategy 'HMO' -Result $result

    Set-ResultView -ResultView $page.ResultView -Status $status -Metrics @(
        @{ Label = 'Gross monthly rent'; Value = Format-Currency $result.GrossMonthlyRent },
        @{ Label = 'Monthly cashflow'; Value = Format-Currency $result.MonthlyCashflow },
        @{ Label = 'Annual cashflow'; Value = Format-Currency $result.AnnualCashflow },
        @{ Label = 'ROI'; Value = Format-Percent $result.ROI },
        @{ Label = 'Gross yield'; Value = Format-Percent $result.GrossYield },
        @{ Label = 'Commercial value'; Value = Format-Currency $result.CommercialValue }
    ) -DetailLines @(
        'HMO BREAKDOWN'
        ''
        ('Stamp duty:             ' + (Format-Currency $result.StampDuty))
        ('Total upfront cost:     ' + (Format-Currency $result.UpfrontTotal))
        ('Gross monthly rent:     ' + (Format-Currency $result.GrossMonthlyRent))
        ('Management fee:         ' + (Format-Currency $result.ManagementFee))
        ('Mortgage amount:        ' + (Format-Currency $result.MortgageAmount))
        ('Monthly mortgage:       ' + (Format-Currency $result.MonthlyMortgage))
        ('Monthly investor cost:  ' + (Format-Currency $result.MonthlyInvestorCost))
        ('Total monthly costs:    ' + (Format-Currency $result.TotalMonthlyCosts))
        ('Monthly cashflow:       ' + (Format-Currency $result.MonthlyCashflow))
        ('Annual cashflow:        ' + (Format-Currency $result.AnnualCashflow))
        ('Cash left in deal:      ' + (Format-Currency $result.CashLeftIn))
        ('ROI:                    ' + (Format-Percent $result.ROI))
        ('Gross yield:            ' + (Format-Percent $result.GrossYield))
        ('Commercial value:       ' + (Format-Currency $result.CommercialValue))
        ('Equity / uplift:        ' + (Format-Currency $result.EquityOrProfit))
    ) -Note 'For England, mandatory HMO licensing often starts at 5+ occupants, but local councils can set wider licensing rules.' -ExecutiveHeadline "$($status.Text): HMO cashflow $(Format-Currency $result.MonthlyCashflow) per month" -ExecutiveBody "Prepared for $(Get-DisplayLeadName). This HMO screens at $(Format-Percent $result.ROI) ROI, $(Format-Percent $result.GrossYield) gross yield, and a commercial value estimate of $(Format-Currency $result.CommercialValue)." -PitchText "HMO deal summary`r`nPrepared for: $(Get-DisplayLeadName)`r`nStatus: $($status.Text)`r`nGross monthly rent: $(Format-Currency $result.GrossMonthlyRent)`r`nMonthly cashflow: $(Format-Currency $result.MonthlyCashflow)`r`nAnnual cashflow: $(Format-Currency $result.AnnualCashflow)`r`nROI: $(Format-Percent $result.ROI)`r`nCommercial value: $(Format-Currency $result.CommercialValue)`r`nSource: https://marcinsakowski.com" -StressText (Get-HmoStressText -Data $data)

    Set-PageExportData -Page $page -Strategy 'HMO' -InputData $data -ResultData @{
        StampDuty = Format-Currency $result.StampDuty
        UpfrontTotal = Format-Currency $result.UpfrontTotal
        GrossMonthlyRent = Format-Currency $result.GrossMonthlyRent
        ManagementFee = Format-Currency $result.ManagementFee
        MortgageAmount = Format-Currency $result.MortgageAmount
        MonthlyMortgage = Format-Currency $result.MonthlyMortgage
        MonthlyInvestorCost = Format-Currency $result.MonthlyInvestorCost
        TotalMonthlyCosts = Format-Currency $result.TotalMonthlyCosts
        MonthlyCashflow = Format-Currency $result.MonthlyCashflow
        AnnualCashflow = Format-Currency $result.AnnualCashflow
        CashLeftIn = Format-Currency $result.CashLeftIn
        ROI = Format-Percent $result.ROI
        GrossYield = Format-Percent $result.GrossYield
        CommercialValue = Format-Currency $result.CommercialValue
        EquityOrProfit = Format-Currency $result.EquityOrProfit
    } -StatusData $status
    $page.ShareText = "HMO deal summary`r`nPrepared for: $(Get-DisplayLeadName)`r`nStatus: $($status.Text)`r`nGross monthly rent: $(Format-Currency $result.GrossMonthlyRent)`r`nMonthly cashflow: $(Format-Currency $result.MonthlyCashflow)`r`nAnnual cashflow: $(Format-Currency $result.AnnualCashflow)`r`nROI: $(Format-Percent $result.ROI)`r`nCommercial value: $(Format-Currency $result.CommercialValue)`r`nSource: https://marcinsakowski.com"

    Write-Log -Message ("HMO calculated: cashflow {0}, ROI {1}." -f (Format-Currency $result.MonthlyCashflow), (Format-Percent $result.ROI))
    Set-BusyLabel -Text 'HMO calculated'
}

$script:ui.NavBTL.Add_Click({ Show-Page 'BTL' })
$script:ui.NavBRRR.Add_Click({ Show-Page 'BRRR' })
$script:ui.NavHMO.Add_Click({ Show-Page 'HMO' })

$script:ui.BTLCalculate.Add_Click({ Invoke-BTLCalculation })
$script:ui.BRRRCalculate.Add_Click({ Invoke-BRRRCalculation })
$script:ui.HMOCalculate.Add_Click({ Invoke-HMOCalculation })
$script:ui.BTLSave.Add_Click({ Save-PageDeal -Page $script:pages.BTL -Strategy 'BTL' })
$script:ui.BRRRSave.Add_Click({ Save-PageDeal -Page $script:pages.BRRR -Strategy 'BRRR' })
$script:ui.HMOSave.Add_Click({ Save-PageDeal -Page $script:pages.HMO -Strategy 'HMO' })
$script:ui.BTLLoad.Add_Click({ Load-PageDeal -Page $script:pages.BTL -Strategy 'BTL' -RecalculateAction { Invoke-BTLCalculation } })
$script:ui.BRRRLoad.Add_Click({ Load-PageDeal -Page $script:pages.BRRR -Strategy 'BRRR' -RecalculateAction { Invoke-BRRRCalculation } })
$script:ui.HMOLoad.Add_Click({ Load-PageDeal -Page $script:pages.HMO -Strategy 'HMO' -RecalculateAction { Invoke-HMOCalculation } })
$script:ui.BTLCopy.Add_Click({ Copy-PageSummary -Page $script:pages.BTL -Strategy 'BTL' })
$script:ui.BRRRCopy.Add_Click({ Copy-PageSummary -Page $script:pages.BRRR -Strategy 'BRRR' })
$script:ui.HMOCopy.Add_Click({ Copy-PageSummary -Page $script:pages.HMO -Strategy 'HMO' })
$script:ui.BTLExport.Add_Click({ Export-PageCsv -Page $script:pages.BTL -Strategy 'BTL' })
$script:ui.BRRRExport.Add_Click({ Export-PageCsv -Page $script:pages.BRRR -Strategy 'BRRR' })
$script:ui.HMOExport.Add_Click({ Export-PageCsv -Page $script:pages.HMO -Strategy 'HMO' })
$script:ui.BTLExportXlsx.Add_Click({ Export-PageXlsx -Page $script:pages.BTL -Strategy 'BTL' })
$script:ui.BRRRExportXlsx.Add_Click({ Export-PageXlsx -Page $script:pages.BRRR -Strategy 'BRRR' })
$script:ui.HMOExportXlsx.Add_Click({ Export-PageXlsx -Page $script:pages.HMO -Strategy 'HMO' })
$script:ui.BTLExportPdf.Add_Click({ Export-PagePdf -Page $script:pages.BTL -Strategy 'BTL' })
$script:ui.BRRRExportPdf.Add_Click({ Export-PagePdf -Page $script:pages.BRRR -Strategy 'BRRR' })
$script:ui.HMOExportPdf.Add_Click({ Export-PagePdf -Page $script:pages.HMO -Strategy 'HMO' })

$script:ui.BTLReset.Add_Click({ Reset-PageDefaults -Page $script:pages.BTL; Write-Log -Message 'BTL defaults restored.' })
$script:ui.BRRRReset.Add_Click({ Reset-PageDefaults -Page $script:pages.BRRR; Write-Log -Message 'BRRR defaults restored.' })
$script:ui.HMOReset.Add_Click({ Reset-PageDefaults -Page $script:pages.HMO; Write-Log -Message 'HMO defaults restored.' })

$openSite = {
    Start-Process 'https://marcinsakowski.com'
}

$script:ui.WebsiteButton.Add_Click($openSite)
$script:ui.BookReviewButton.Add_Click($openSite)
$script:ui.WebsiteLink.Add_RequestNavigate({
    param($sender, $e)
    Start-Process $e.Uri.AbsoluteUri
    $e.Handled = $true
})

try {
    Reset-PageDefaults -Page $script:pages.BTL
    Reset-PageDefaults -Page $script:pages.BRRR
    Reset-PageDefaults -Page $script:pages.HMO

    Show-Page 'BTL'
    Write-Log -Message 'Calculator loaded.'
    Invoke-BTLCalculation

    [void]$window.ShowDialog()
}
catch {
    [System.Windows.MessageBox]::Show(
        "The calculator failed during startup.`n`n$($_.Exception.Message)",
        'UK Property Investment Calculator',
        'OK',
        'Error'
    ) | Out-Null
    throw
}
