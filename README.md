<div align="center">

## access 97 password recovery


</div>

### Description

Forgot your access 97 password?

Never fear this function will

hack the access 97 local file

password
 
### More Info
 
an access 97 file with local password set

I hope this code will end the commercial password recoverers

the password


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Tom Pickles](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tom-pickles.md)
**Level**          |Unknown
**User Rating**    |4.3 (176 globes from 41 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tom-pickles-access-97-password-recovery__1-4044/archive/master.zip)

### API Declarations

```
Original code translated direct from c snippet by Casper Ridley
From the c code at http://www.nmrc.org/
Turned into this nifty function by me =)
```


### Source Code

```
Function AccessPassword(Byval Filename As string) as string
Dim MaxSize, NextChar, MyChar, secretpos,TempPwd
Dim secret(13)
secret(0) = (&H86)
secret(1) = (&HFB)
secret(2) = (&HEC)
secret(3) = (&H37)
secret(4) = (&H5D)
secret(5) = (&H44)
secret(6) = (&H9C)
secret(7) = (&HFA)
secret(8) = (&HC6)
secret(9) = (&H5E)
secret(10) = (&H28)
secret(11) = (&HE6)
secret(12) = (&H13)
secretpos = 0
Open Filename For Input As #1  ' Open file for input.
For NextChar = 67 To 79 Step 1 'Read in Encrypted Password
 Seek #1, NextChar      ' Set position.
 MyChar = Input(1, #1)    ' Read character.
 TempPwd = TempPwd & Chr(Asc(MyChar) Xor secret(secretpos)) 'Decrypt using Xor
 secretpos = secretpos + 1  'increment pointer
Next NextChar
Close #1  ' Close file.
AccessPassword = TempPwd
End Function
```

