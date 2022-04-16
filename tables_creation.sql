
IF EXISTS
( SELECT 1
	FROM sysobjects o
	WHERE	(o.[name] = 'pr_rmv_table')
	AND	(OBJECTPROPERTY(o.[ID],'IsProcedure')=1)
)
BEGIN
	DROP PROCEDURE pr_rmv_table
		
END

GO

CREATE PROCEDURE [dbo].pr_rmv_table
(	@table_name nvarchar(100)
)
AS
/* Procedura sprawdza czy istnieje w bazie tabela @table_name
** Jak tak to usuwa ją
*/
	DECLARE @stmt nvarchar(1000)

	IF EXISTS 
	( SELECT 1
		FROM sysobjects o
		WHERE	(o.[name] = @table_name)
		AND	(OBJECTPROPERTY(o.[ID],'IsUserTable')=1)
	)
	BEGIN
		SET @stmt = 'DROP TABLE ' + @table_name
		EXECUTE sp_executeSQL @stmt = @stmt
	END
GO

EXEC pr_rmv_table @table_name='returns_details'
EXEC pr_rmv_table @table_name='order_items'
EXEC pr_rmv_table @table_name='orders'
EXEC pr_rmv_table @table_name='clients'
EXEC pr_rmv_table @table_name='storage'
EXEC pr_rmv_table @table_name='books '



create table [dbo].books 
(	isbn 	numeric(13) 	not null identity constraint pk_books primary key
,	title 		varchar(100) 	not null
,	author		varchar(50)		not null
,	published_year		varchar(4)		not null
,	descr		varchar(500)	not null
,	publisher		varchar(50)		not null
,	price		money			not null
,	genre		varchar(50)
)
GO

create table [dbo].storage
(	isbn	numeric(13) not null constraint fk_storage__books foreign key references books(isbn)
,	storage_count	int	not null
)
GO

create table [dbo].clients
(	id_client int not null	identity constraint pk_clients primary key
,	name_client varchar(30) not null
,	surname_client varchar(30) not null
,	client_email	varchar(40) not null
,	client_addres varchar(60)
,	client_phone	numeric(12)
)
GO

create table [dbo].orders(
	id_order int not null identity constraint pk_orders primary key
,	id_client	int not null constraint fk_orders__clients foreign key references clients(id_client)
,	order_date	datetime	not null
,	order_addres	varchar(60)		not null
,	order_phone	numeric(12)		not null
)
GO

create table [dbo].order_items
(	id_order_item	int 		not null identity constraint pk_order_item primary key
,	isbn	numeric(13) not null constraint fk_order_items__books foreign key references books(isbn)
,	item_count int not null
,	id_order	int not null constraint fk_order_items__orders foreign key references orders(id_order)
)
GO

create table [dbo].returns_details
(	id_order_item	int not null constraint fk_returns__order_items	 foreign key references	order_items(id_order_item)
,	return_date		datetime	not null
)
GO
