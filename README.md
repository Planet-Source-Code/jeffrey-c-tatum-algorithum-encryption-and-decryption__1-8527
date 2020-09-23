<div align="center">

## Algorithum Encryption and Decryption


</div>

### Description

This code will take text, and encrypt it in Algorithum Encryption. What it does is randomly create 4 keys, each one 1 char long. No password needed to encrypt the text, it creates its own and stores it within the encrypted text. Virtualy impossible to crack because the output encryption is rarely the same. For example, a string that says "test" will be 8 charachters long, and almost never the same. Or a string that says "testing" will be 11 charachters long. The key to encrypt and decrypt is stored at the beginning and the end of the output making it almost impossible to crack.
 
### More Info
 
To see how this code works, create 3 text box's. Text1, Text2, and Text3.

In Text1_Change, put:

text2 = TEncrypt(text1)

In Text2_Change, put:

text3 = TDecrypt(text2)

Put nothing in Text3_Change.

What this will show you is the text you type in Text1, will get encrypted into Text2. It will then Decrypt itself and show you in Text3.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jeffrey C\. Tatum](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jeffrey-c-tatum.md)
**Level**          |Intermediate
**User Rating**    |4.3 (81 globes from 19 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jeffrey-c-tatum-algorithum-encryption-and-decryption__1-8527/archive/master.zip)





### Source Code

```
Function TEncrypt (iString)
On Error GoTo uhoh
Q = ""
a = randomnumber(9) + 32
b = randomnumber(9) + 32
c = randomnumber(9) + 32
d = randomnumber(9) + 32
Q = Chr(a) & Chr(c) & Chr(b)
e = 1
For x = 1 To Len(iString)
f = Mid(iString, x, 1)
  If e = 1 Then Q = Q & Chr(Asc(f) + a)
  If e = 2 Then Q = Q & Chr(Asc(f) + c)
  If e = 3 Then Q = Q & Chr(Asc(f) + b)
  If e = 4 Then Q = Q & Chr(Asc(f) + d)
e = e + 1
If e > 4 Then e = 1
Next x
Q = Q & Chr(d)
TEncrypt = Q
Exit Function
uhoh:
TEncrypt = "Error: Invalid text to Encrypt"
Exit Function
End Function
Function TDecrypt (iString)
On Error GoTo uhohs
Q = ""
zz = Left(iString, 3)
a = Left(zz, 1)
b = Mid(zz, 2, 1)
c = Mid(zz, 3, 1)
d = Right(iString, 1)
a = Int(Asc(a)) 'key 1
b = Int(Asc(b)) 'key 2
c = Int(Asc(c)) 'key 3
d = Int(Asc(d)) 'key 4
txt = Left(iString, Len(iString) - 1)
txt2 = Mid(txt, 4, Len(txt)) 'encrypted text
e = 1
For x = 1 To Len(txt2)
f = Mid(txt2, x, 1)
  If e = 1 Then Q = Q & Chr(Asc(f) - a)
  If e = 2 Then Q = Q & Chr(Asc(f) - b)
  If e = 3 Then Q = Q & Chr(Asc(f) - c)
  If e = 4 Then Q = Q & Chr(Asc(f) - d)
e = e + 1
If e > 4 Then e = 1
Next x
TDecrypt = Q
Exit Function
uhohs:
TDecrypt = "Error: Invalid text to Decrypt"
Exit Function
End Function
Function randomnumber (finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function
```

