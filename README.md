# 🗓️ Create Outlook Lunch Blockers

This PowerShell script automatically generates "busy" calendar appointments in Microsoft Outlook during lunchtime hours on weekdays within a specified date range.

## 📋 Features

- Creates **private, busy** appointments in your Outlook calendar
- Schedules time blocks between **11:00 AM and 12:15 PM**
- Randomized start time (in 15-minute intervals)
- Randomized duration: **90 or 120 minutes**
- Weekdays only (Monday–Friday)
- Assigns the category: `Lunch`
- No reminder set
- High importance

## 📅 Default Configuration

- **Date Range**: December 1–23, 2025  
- **Days**: Monday through Friday  
- **Start Times**: Between 11:00 and 12:15  
- **Durations**: 90 or 120 minutes  
- **Outlook Calendar**: Uses your default Outlook profile via COM automation

## 🚀 How to Use

1. Open **PowerShell** as a user who has access to Outlook.
2. Make sure Outlook is installed and configured on your system.
3. Run the script:

   ```powershell
   .\Create_Outlook_Lunch_Blockers.ps1
   ```

4. The script will silently create appointments in your Outlook calendar based on the defined rules.

## 🔧 Customization

You can modify the following parameters in the script:

| Variable      | Description                        |
|---------------|------------------------------------|
| `$StartDate`  | Beginning of scheduling window     |
| `$EndDate`    | End of scheduling window           |
| `$TimeSlots`  | Start time options (in decimal)    |
| `$Durations`  | Duration options in minutes        |

### Example Time Slot Format
```powershell
$TimeSlots = @(11, 11.25, 11.5, 11.75, 12, 12.25)
```
Decimal values represent time as follows:  
- `11.25` = 11:15  
- `11.5` = 11:30  
- `11.75` = 11:45  
- `12.25` = 12:15

## ⚠️ Requirements

- Windows OS
- Microsoft Outlook installed and configured
- PowerShell 5.1+ (default on most Windows systems)

## ✅ Example Appointment

- **Subject**: `Blocker (1.5 Std)`
- **Time**: 11:15 AM – 12:45 PM
- **Status**: Busy
- **Sensitivity**: Private
- **Category**: Lunch
- **Importance**: High

## 🧪 Development Tips

- To test without saving to your calendar, comment out this line:
  ```powershell
  $Appointment.Save()
  ```
  and replace with:
  ```powershell
  Write-Output "Would create: $StartTime ($Duration minutes)"
  ```

---

## ✍️ Author

Created by GentleBerry – feel free to contribute or suggest improvements!
