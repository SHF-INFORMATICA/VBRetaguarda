use megasim
  
  update cep set Cep = '0' + CEP
  where len(cep) = 7
  
  
  update ENDERECO set Cep = '0' + CEP
  where len(cep) = 7
  
delete from cep
where LEN(cep) = 6
  
delete from ENDERECO
where LEN(cep) = 6