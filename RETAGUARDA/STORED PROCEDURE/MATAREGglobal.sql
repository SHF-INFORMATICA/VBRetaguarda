delete mfi010
FROM   MFA010 INNER JOIN
       MFI010 ON MFA010.MFASEQUENCIA = MFI010.MFISEQUEN
WHERE MFAEMISSAO < '01/07/2018 00:00:00' 

delete se1010 WHERE E1_EMISSAO < '01/07/2018 00:00:00' 

delete mfa010 WHERE MFAEMISSAO < '01/07/2018 00:00:00' 