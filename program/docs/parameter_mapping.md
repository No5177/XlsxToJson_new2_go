# Parameter Mapping Documentation

## File Structure
- Input: Excel file with device parameters (No.1~95, including 8A, 8B)
- Output: FWSpecification.json following spec format

## Device-Specific Analysis
Comparing TPT-B1RS0505A vs TPT-B4R1010A:

1. Voltage Ranges
   - TPT-B1RS0505A: VSETmax=5000, VSETmin=100
   - TPT-B4R1010A: VSETmax=20000, VSETmin=400
   
2. Current Ranges
   - TPT-B1RS0505A: CISETmax=5000, CISETmin=2.5
   - TPT-B4R1010A: CISETmax=10000, CISETmin=5

3. Power Ranges
   - TPT-B1RS0505A: PSETmax=25, PSETmin=0.0125
   - TPT-B4R1010A: PSETmax=200, PSETmin=0.1

Key observations:
- Different device models have different parameter ranges
- Ratios between models are consistent (approximately 4x for TPT-B4R1010A vs TPT-B1RS0505A)
- Resolution values remain constant across devices

## JSON Structure Analysis
Based on FWSpecification_spec.json and example files:

### Core Parameter Categories
1. Voltage Settings (VSET*)
   - VSETmax, VSETmin: Voltage range limits
   - VSETreso: Voltage resolution
   - VSETACCmax, VSETACCmin: Voltage accuracy range

2. Current Settings (CISET*, DISET*)
   - CISETmax, CISETmin: Charge current range
   - DISETmax, DISETmin: Discharge current range
   - CISETreso, DISETreso: Current resolutions

3. Power Settings (PSET*, CPSET*)
   - PSETmax, PSETmin: Power range limits
   - CPSETmax, CPSETmin: Constant power range
   - PSETreso: Power resolution

4. Time Settings
   - Timesetmax, Timesetmin: Time range limits
   - RecordTimesetmax, RecordTimesetmin: Recording time range
   - Timesetreso, RecordTimesetreso: Time resolutions

5. Mean Value Parameters (VMEAN*, CIMEAN*, DIMEAN*, PMEAN*)
   - Range limits (max/min)
   - Resolution values (reso)
   - Accuracy ranges (ACCmax/ACCmin)

6. Protection Parameters
   - OVP: Over-voltage protection
   - UVP: Under-voltage protection
   - COCP: Charge over-current protection
   - DOCP: Discharge over-current protection
   - OTP: Over-temperature protection

## Excel to JSON Mapping
Parameters in Excel file (No.1~95, including 8A, 8B) map to JSON fields as follows:
1. VSETmax -> "VSETmax"
2. VSETmin -> "VSETmin"
3. CISETmax -> "CISETmax"
4. CISETmin -> "CISETmin"
...etc.

## Device Name Validation
Format pattern: TPT-B[number/letter combination]
Examples: 
- TPT-B1RS0505A
- TPT-B4R1010A

## Parameter Relationships and Validation Rules

1. Parameter Dependencies
   - Power values (PSETmax) are derived from voltage and current: P = V * I
   - Mean values (VMEAN, CIMEAN, etc.) have wider ranges than their SET counterparts
   - Protection values (OVP, COCP, etc.) are slightly higher than their SET counterparts

2. Resolution Rules
   - VSETreso, CISETreso, DISETreso: Always 1
   - PSETreso: Always 0.001
   - Timesetreso: Always 100
   - TSETreso, OTPreso: Always 0.1

3. Validation Rules
   - All min values must be less than their corresponding max values
   - Protection values must be higher than operational values:
     * OVPmax > VSETmax
     * COCPmax > CISETmax
     * DOCPmax > DISETmax
   - Mean ranges must encompass operational ranges:
     * VMEANmax > VSETmax
     * CIMEANmax > CISETmax
     * DIMEANmax > DISETmax

4. Device-Specific Rules
   - Model number indicates capacity:
     * B4R series: Higher capacity (VSETmax=20000, CISETmax=10000)
     * B1RS series: Lower capacity (VSETmax=5000, CISETmax=5000)
   - Resolution values remain constant across all models
   - Ratio between min/max values stays consistent within series

5. Value Ranges (Base Ranges)
   - Voltage: 0-20000
   - Current: 0-10000
   - Power: 0-200
   - Temperature: 0-130
   - Time: 100-3599999900
   - Alarm Delay: 100-60000
