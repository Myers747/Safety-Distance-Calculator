# Safety Distance Calculator

A Windows Forms application to calculate minimum safe distances for automation robotics, based on ISO 13855 guidelines.

Developed in Visual Basic .NET using the EPPlus library for Excel file interaction.

---

## 📋 Features

- Calculates minimum safe distances based on actuator and light curtain timing values
- Dynamic UI for selecting Safety Controllers, PLCs, Light Curtains, and Actuators
- Live overview of timing (T) and offset (C) contributions
- Manual calculation trigger to ensure user-reviewed inputs
- Enforces a minimum 4-inch (101.6mm) safe distance
- Designed for deployment in network-shared drive environments

---

##  Project Structure

```
/Release
    Distance_UI.exe
    Safe_Distance_Data.xlsx
    EPPlus.dll
/Source
    Distance_UI.sln
    Distance_UI/
README.md
LICENSE
```

---

##  Getting Started

1. Download the `/Release` folder contents.
2. Place the `Safe_Distance_Data.xlsx` alongside the executable.
3. Run `Distance_UI.exe`.
4. (Optional) Open `Safe_Distance_Data.xlsx` directly to edit actuator and safety device settings.

---

##  Requirements

- Windows 10 or newer
- .NET Framework 4.8 or higher
- Excel file named `Safe_Distance_Data.xlsx` stored alongside the application

---

##  License

This project is licensed under the MIT License. See the `LICENSE` file for details.

---

##  Author

Built with ❤️ by **Tamara Myers**  
for practical automation safety solutions.

---
