# Security module
Cross-platform encryption system (SQL / VBA / Python)

This folder contains cross-plataform encryption utilities designed to operate consistently across:
- ** SQL Server (T-SQL)**
- ** VBA / Microsoft Access**
- ** Python**
- The goal of this design is deterministic compatibility
- data encrypted in one environment can be decrypted correctly in the others.

- ## Highlights
- Custom cipher with randomized components
- No dependency on external crypto libraries
- Safe for use in mixed-technology stacks (Access - SQL Server - Python)

- ## Cross-Platform Compatibility
- The same encrypted value produced by SQL Server can be:
- decrypted by VBA
- decrypted by Python
- and verified again by SQL Server

- This allows encrypted values to move freely between systems without data loss.

- ## Contents
- Each language repository contains it's own implementations of the same algorithm.

- ## Example
- If input string = "myString", multiple outputs of the same input produce different results:
- Immediate Window:
- ? CipherString("mystring")  
5050667C666A665864647C646A5841667750597C4569FF44EB75937B29A948A15DF77E5C45459AF9F193D2A35E49F07EFCE647CA68ABF8B8FBB42EF37FC9EAE4456EA1C83A284367BE75AE92305CBF594C302380AF33674F7582FFD4E56688729C8CB26EF88DC5FEE83DCF48A683D58962EA5E664583264A7977922C407536D7

- ? CipherString("mystring")
5E4D5E487D51515748485E486A594D57517D5F5E7A6D6D42DAEFF8D7913EE1CFE873724EDCB8775E8B884AB35BA89272B88D63A7422DE55961AE802779F125A3DAFBD7AFA73961DC7B76C795FDB050D75BA65665634A6CAFE792FB5966B58849ABC4A06F50C8F27E9785376EB97C544A65CBB267C8E44682AADCD93141408BBE  

- ? CipherString("mystring")
7D67665E66667D5B5E5E4E5E667D675B4F78484E5A6961C2E87D829AE85DA15C9D2679608A4DA280E4C42072C09229737A7779EA8E6D39594A53D1907FE05AD85F3CE6F05A9FB78B3C72B98232E034B07D58FC2862B560D69B212F5D7DE45B8DA53B846A8776E9F1CE68B34FA169BC4467C33424AC63BF753E5CA4AEF086A165  

