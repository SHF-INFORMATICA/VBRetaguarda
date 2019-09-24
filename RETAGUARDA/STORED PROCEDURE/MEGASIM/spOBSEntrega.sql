--create procedure spOBSEntrega  @Acao int,@OBSENTREGA_ID bigint,@ENTREGA_ID bigint,@OBS nvarchar(MAX),@DT_CAD datetime
alter procedure spOBSEntrega  @Acao int,@OBSENTREGA_ID bigint,@ENTREGA_ID bigint,@OBS nvarchar(MAX),@DT_CAD datetime
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
			SET @OBSEntrega_ID = (select max(OBSEntrega_id) from OBSEntrega) + 1 

			IF @OBSEntrega_ID IS NULL 
				SET @OBSEntrega_ID = 1
			IF @OBSEntrega_ID <= 0 
				SET @OBSEntrega_ID = 1

			INSERT INTO OBSEntrega (OBSEntrega_ID, ENTREGA_ID, OBS, DT_CAD)
			Values (@OBSEntrega_ID,@ENTREGA_ID,@OBS,getdate())

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update OBSEntrega SET OBS = @OBS
			where OBSEntrega_id = @OBSEntrega_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 3) begin        
		BEGIN TRANSACTION
			delete from OBSEntrega where Entrega_id = @Entrega_ID     

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	begin raiserror('Erro, não executado, spOBSEntrega',14,1)     
	End  
End 