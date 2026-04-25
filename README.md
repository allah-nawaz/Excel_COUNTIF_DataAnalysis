# Excel COUNTIF Questions For Data Analyst Interview Preparation

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

## COUNTIF Practice Questions & Answers

---

### 1. Count students from Karachi
```excel
=COUNTIF(C2:C7,"Karachi")
```

### 2. Count students with marks greater than or equal to 60
```excel
=COUNTIF(B2:B7,">=60")
```
### 3. Count students with marks less than or equal to 60
```excel
=COUNTIF(B2:B7,"<=60")
```
### 4. Count students who are NOT from Lahore
```excel
=COUNTIF(C2:C7,"<>Lahore")
```
### 5. Count students whose name is Ali
```excel
=COUNTIF(A2:A7,"Ali")
```
### 6. Count how many times marks are exactly 60
```excel
=COUNTIF(B2:B7,60)
```
### 7. Count students whose marks are between 50 and 80
```excel
=COUNTIFS(B2:B7,">50",B2:B7,"<=80")
```
### 8. Count students whose city starts with "K"
```excel
=COUNTIF(C2:C7,"K*")
```
### 9. Count students whose name has exactly 4 letters
```excel
=COUNTIF(A2:A7,"????")
```
### 10. Count students based on value in cell D14 (city match)
```excel
=COUNTIF(C2:C7,D14)
```
### 11. Count marks greater than value in cell D14
```excel
=COUNTIF(B2:B7,">"&D14)
```
### 12. Count students from Islamabad
```excel
=COUNTIF(C2:C7,"Islamabad")
```
### 13. Count students with marks less than 40
```excel
=COUNTIF(B2:B7,"<40")
```
### 14. Count cities ending with "i"
```excel
=COUNTIF(C2:C7,"*i")
```
### 15. Count names starting with "A"
```excel
=COUNTIF(A2:A7,"A*")
```

### Imp Notes From this Project:
COUNTIF = single condition counting <br>
COUNTIFS = multiple conditions <br>
Wildcards: * and ?

