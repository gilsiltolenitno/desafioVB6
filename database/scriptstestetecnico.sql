--CRIAÇÃO DA BASE DE TESTE
CREATE DATABASE XYZAdmCardDB;

GO XYZAdmCardDB;

--TABELA DE CLIENTES
CREATE TABLE Clientes
  (
    IdCliente INT IDENTITY(1,1) PRIMARY KEY,
    NomeCliente VARCHAR(100) NOT NULL,
    NumeroCartao VARCHAR(20) NOT NULL UNIQUE,
    DataCadastro DATETIME DEFAULT GETDATE()
  );

--TABELA DE TRANSACOES
CREATE TABLE Transacoes
  (
    IdTransacao INT IDENTITY(1,1) PRIMARY KEY,
    NumeroCartao VARCHAR(20) NOT NULL,
    ValorTransacao DECIMAL(10,2) NOT NULL,
    DataTransacao DATETIME NOT NULL,
    Descricao VARCHAR(255),
    
    CONSTRAINT FK_Transacoes_Clientes
    FOREIGN KEY (NumeroCartao)
    REFERENCES Clientes(NumeroCartao)
);

ALTER TABLE Transacoes
ADD IdCliente INT NULL;


ALTER TABLE Transacoes
ADD CONSTRAINT FK_Transacoes_Clientes_Id
FOREIGN KEY (IdCliente)
REFERENCES Clientes(IdCliente)

---CARGA INICIAL DE CLIENTES
INSERT INTO Clientes (NomeCliente, NumeroCartao) VALUES
  ('GILBERTO TOLENTINO','4000000000000001'),
  ('FERNANDA BARROS','4000000000000002'),
  ('MARIANA TOLENTINO', '4000000000000003'),
  ('JULIANO ANTUNES','4000000000000004');

--Procedure para calcular Total de transações
CREATE PROCEDURE PRTotalTransacoesPorCartao

  @DataInicial Date,
  @DataFinal   Date

  AS 
  BEGIN
  SET NOCOUNT ON;
    SELECT NumeroCartao,
      COUNT(*) AS QTDTransacoes,
      SUM (ValorTransacao) AS TOTValor
    FROM TRANSACOES
    WHERE DataTransacao BETWEEN @DataInicial AND @DataFinal
    GROUP BY NumeroCartao
    ORDER BY NumeroCartao

END

--Função para categorizar transações
CREATE FUNCTION fn_CATegoriaTransacoes
(
  @Valor_Transacao DECIMAL(10,2)
)
  RETURNS VARCHAR(10)
  AS
  BEGIN
    RETURN
      CASE
       WHEN @Valor_Transacao > 1000 THEN 'ALTA'
       WHEN @Valor_Transacao BETWEEN 500 AND 1000 THEN 'MÉDIA'
       WHEN @Valor_Transacao < 500 THEN 'BAIXA'
      END
  END

  SELECT 
    NumeroCartao,
    ValorTransacao,
    DBO.fn_CATegoriaTransacoes(ValorTransacao)
  FROM Transacoes

-Criação da view de clientes e transações
  CREATE VIEW vw_TransacoesClientes
  AS
    SELECT 
      CLI.IdCliente       AS Codigo,
      CLI.NomeCliente     AS Nome,
      TRA.NumeroCartao    AS Cartao,
      TRA.ValorTransacao  AS Valor,
      TRA.DataTransacao   AS DATA,
      dbo.fn_CATegoriaTransacoes(TRA.ValorTransacao) AS CATegoria
    FROM Transacoes TRA
    JOIN Clientes CLI
    ON CLI.NumeroCartao = TRA.NumeroCartao





