# 📊 Excel COUNTIF Practice For Data Analyst Interview Preparation

This project shows practical use of Excel `COUNTIF` through interview type questions.

---

## Dataset Used

| Name  | Marks | City       |
|-------|------|------------|
| Ali   | 45   | Lahore     |
| Sara  | 78   | Karachi    |
| Ahmed | 90   | Lahore     |
| Zain  | 60   | Islamabad  |
| Hina  | 55   | Lahore     |
| Bilal | 30   | Karachi    |

---

## 📊 COUNTIF Practice Questions & Answers

---

### 1. Count students from Karachi
```excel
=COUNTIF(C2:C7,"Karachi")

---

### 1. Count students with marks greater than or equal to 60
=COUNTIF(B2:B7,">=60")
Count students with marks less than or equal to 60
=COUNTIF(B2:B7,"<=60")
Count students who are NOT from Lahore
=COUNTIF(C2:C7,"<>Lahore")
Count students whose name is Ali
=COUNTIF(A2:A7,"Ali")
Count how many times marks are exactly 60
=COUNTIF(B2:B7,60)
Count students whose marks are between 50 and 80
=COUNTIFS(B2:B7,">50",B2:B7,"<=80")
Count students whose city starts with "K"
=COUNTIF(C2:C7,"K*")
Count students whose name has exactly 4 letters
=COUNTIF(A2:A7,"????")
Count students based on value in cell D14 (city match)
=COUNTIF(C2:C7,D14)
Count marks greater than value in cell D14
=COUNTIF(B2:B7,">"&D14)
Count students from Islamabad
=COUNTIF(C2:C7,"Islamabad")
Count students with marks less than 40
=COUNTIF(B2:B7,"<40")
Count cities ending with "i"
=COUNTIF(C2:C7,"*i")
Count names starting with "A"
=COUNTIF(A2:A7,"A*")
🧠 Key Learning Points
COUNTIF = single condition counting
COUNTIFS = multiple conditions
Wildcards: * and ?
Dynamic criteria using "&cell_reference"
