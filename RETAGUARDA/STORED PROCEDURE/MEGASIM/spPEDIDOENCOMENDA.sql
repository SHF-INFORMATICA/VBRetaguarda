alter procedure spPEDIDOENCOMENDA	@Acao int,
									@PEDIDOENCOMENDA_ID bigint,
									@PEDIDO_ID bigint,
									@USUARIO_ID bigint,
									@VLR_TX_ENTREGA float
as  begin
	if (@Acao = 1) begin
		BEGIN TRANSACTION
		    SET @PEDIDOENCOMENDA_ID = (select max(PEDIDOENCOMENDA_id) from PEDIDOENCOMENDA) + 1

			IF @PEDIDOENCOMENDA_ID IS NULL 
				SET @PEDIDOENCOMENDA_ID = 1
			IF @PEDIDOENCOMENDA_ID <= 0 
				SET @PEDIDOENCOMENDA_ID = 1

			INSERT INTO PEDIDOENCOMENDA (PEDIDOENCOMENDA_ID, PEDIDO_ID, DT_RECEBE, USUARIO_ID, VLR_TX_ENTREGA)
			Values (@PEDIDOENCOMENDA_ID,@PEDIDO_ID,getdate(),@USUARIO_ID,@VLR_TX_ENTREGA)

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End     
	else if (@Acao = 2) begin        
		BEGIN TRANSACTION
			update PEDIDOENCOMENDA SET DT_RECEBE = GETDATE(), USUARIO_ID = @USUARIO_ID, VLR_TX_ENTREGA = @VLR_TX_ENTREGA
			where PEDIDO_ID = @PEDIDO_ID

		IF @@ERROR <> 0 
		   ROLLBACK 
		   ELSE
		      COMMIT
	End
	else if (@Acao = 3) begin
		BEGIN TRANSACTION
			delete from PEDIDOENCOMENDA where PEDIDO_ID = @PEDIDO_ID

		IF @@ERROR <> 0
		   ROLLBACK
		   ELSE
		      COMMIT
	End
	begin raiserror('Erro, não executado, spPEDIDOENCOMENDA',14,1)
	End
End