
create database MOVISTAR

restore database MOVISTAR from disk = 'C:\MOVISTAR'

use MOVISTAR

create table Vendedor
(
Num_Cedula char(16) check(Num_Cedula like '[0-9][0-9][0-9]-[0-9][0-9][0-9][0-9][0-9][0-9]-[0-9][0-9][0-9][0-9][A-Z]') primary key not null,
Sexo nvarchar(15) not null,
Edad int null,
I_Nombre nvarchar(15) not null,
II_Nombre nvarchar(15) not null,
I_Apellido nvarchar(15) not null,
II_Apellido nvarchar(15) not null,
Estado_Civil nvarchar(15) not null,
Direccion nvarchar(70) not null
)

create table Venta_Contado
(
Id_Venta int identity(1,1) primary key not null,
Fecha_Venta smalldatetime not null,
Total_C$ money null,
Total_Comision_Vendedor money null,
Total_Comision_Arqueador money null,
Total_Comision_Administrador money null,
Num_Cedula char(16) foreign key references Vendedor(Num_Cedula) not null
)

create table Detalle_Venta_Contado
(
Id_Venta int identity(1,1) primary key not null,
Tipo char(2) not null,
Disponible int not null,
Cantidad int not null,
Descuento float not null,
Precio money null,
SubTotal money null,
Devolucion int null,
Comision_Vendedor money null,
Comision_Arqueador money null,
Comision_Administrador money null,
Fecha_Venta smalldatetime not null, 
Num_Cedula char(16) foreign key references Vendedor(Num_Cedula) not null
)

create table Usuario
(
Id_Usuario int identity(1,1) primary key not null,
Nombre nvarchar(35) not null,
Contraseña nvarchar(35) not null,
Tipo nvarchar(35) not null
)

-- Procedimientos, Funciones y Disparadores

create function Calcular_SubTotal(@Tipo char(3), @Cantidad int)
returns money
begin
 declare @precio as int
 declare @st as money
 set @precio = (select cast(@Tipo as int)) 
 set @st = (@precio * @Cantidad)
 return @st
end

create function Obtener_Precio(@Tipo char(3),@Descuento float)
returns money
begin
 declare @entero as int
 declare @porcentaje as float
 declare @valor as money
 declare @precio as money
 set @entero = (select cast(@Tipo as int))
 set @porcentaje = @Descuento / 100
 set @valor = @entero * @porcentaje
 set @precio = @entero - @valor
 return @precio
end

create function Obtener_SubTotal(@Cantidad int,@Precio money)
returns money
begin
 declare @st as money
 set @st = (@Cantidad * @Precio)
 return @st
end

create function Obtener_Devolucion(@Disponible int,@Cantidad int)
returns int
begin
 declare @dev as int
 set @dev = @Disponible - @Cantidad
 return @dev
end

create procedure Ingresar_Vendedor
@Num_Cedula char(16),
@Sexo nvarchar(15),
@I_Nombre nvarchar(15),
@II_Nombre nvarchar(15),
@I_Apellido nvarchar(15),
@II_Apellido nvarchar(15),
@Estado_Civil nvarchar(15),
@Direccion nvarchar(70)
as
insert into Vendedor values(@Num_Cedula,@Sexo,dbo.Calcular_Edad(@Num_Cedula,getdate()),@I_Nombre,@II_Nombre,@I_Apellido,@II_Apellido,@Estado_Civil,@Direccion)

sp_bindrule 'validar','Vendedor.Edad'

create function Calcular_Edad(@Num_Cedula char(16),@fecha datetime)
returns int
begin
 declare @cedula as nvarchar(16)
 declare @ced as nvarchar(6)
 declare @año as char(2)
 declare @mes as char(2)
 declare @dia as char(2)
 declare @naci as datetime
 declare @dif as datetime
 declare @string as nvarchar(2)
 declare @convert as nvarchar(30)
 declare @edad as int 
 set @cedula = @Num_Cedula
 set @dia = (select substring(@cedula,5,6))
 set @mes = (select substring(@cedula,7,8))
 set @año = (select substring(@cedula,9,10))
 set @ced = @año + @mes + @dia
 set @naci = @ced
 set @dif = @fecha - @naci
 set @convert = (select cast(@dif as nvarchar))
 set @string = (select substring(@convert,10,11))
 set @edad = (select cast(@string as int))
 return @edad
end

create procedure Ingresar_Venta_Contado
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
insert into Venta_Contado values(@Fecha_Venta,0,0,0,0,@Num_Cedula)

sp_bindrule 'validar','Venta_Contado.Total_C$'
sp_bindrule 'validar','Venta_Contado.Total_Comision_Vendedor'
sp_bindrule 'validar','Venta_Contado.Total_Comision_Arqueador'
sp_bindrule 'validar','Venta_Contado.Total_Comision_Administrador'
sp_bindrule 'validar','Detalle_Venta_Contado.Disponible'
sp_bindrule 'validar','Detalle_Venta_Contado.Cantidad'
sp_bindrule 'validar','Detalle_Venta_Contado.Descuento'
sp_bindrule 'validar','Detalle_Venta_Contado.Devolucion'
sp_bindrule 'validar','Detalle_Venta_Contado.Comision_Vendedor'
sp_bindrule 'validar','Detalle_Venta_Contado.Comision_Arqueador'
sp_bindrule 'validar','Detalle_Venta_Contado.Comision_Administrador'
sp_bindrule 'validar','Detalle_Venta_Contado.SubTotal'

create procedure Ingresar_Detalle_Venta_Contado
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
insert into Detalle_Venta_Contado values(@Tipo,@Disponible,@Cantidad,@Descuento,dbo.Obtener_Precio(@Tipo,@Descuento),dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),dbo.Obtener_Devolucion(@Disponible,@Cantidad),dbo.Calcular_Comision_Vendedor(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),dbo.Calcular_Comision_Arqueador(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),dbo.Calcular_Comision_Administrador(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),@Fecha_Venta,@Num_Cedula)

create function Calcular_Comision_Vendedor(@Descuento float,@SubTotal money)
returns money
begin
 declare @sal as money
 set @sal = @SubTotal * 0.0125
 return @sal
end

create function Calcular_Comision_Arqueador(@Descuento float,@SubTotal money)
returns money
begin
 declare @sal as money
 set @sal = @SubTotal * 0.0025
 return @sal
end

create function Calcular_Comision_Administrador(@Descuento float,@SubTotal money)
returns money
begin
 declare @sal as money
 set @sal = @SubTotal * 0.0050
 return @sal
end

create trigger Actualizar_Venta_Contado
on Detalle_Venta_Contado after insert
as
update Venta_Contado set Total_C$ = Total_C$ + (select SubTotal from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Contado set Total_Comision_Vendedor = Total_Comision_Vendedor + (select Comision_Vendedor from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Contado set Total_Comision_Administrador = Total_Comision_Administrador + (select Comision_Administrador from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Contado set Total_Comision_Arqueador = Total_Comision_Arqueador + (select Comision_Arqueador from inserted)
from inserted i,Venta_Contado v
where v.Fecha_Venta = i.Fecha_Venta and v.Num_Cedula = i.Num_Cedula

-- Procedimientos de modificación

create procedure Modificar_Vendedor
@Num_Cedula char(16),
@Sexo nvarchar(15),
@I_Nombre nvarchar(15),
@II_Nombre nvarchar(15),
@I_Apellido nvarchar(15),
@II_Apellido nvarchar(15),
@Estado_Civil nvarchar(15),
@Direccion nvarchar(70)
as
update Vendedor set Sexo = @Sexo,
Edad = dbo.Calcular_Edad(@Num_Cedula,getdate()),
I_Nombre = @I_Nombre,
II_Nombre = @II_Nombre,
I_Apellido = @I_Apellido,
II_Apellido = @II_Apellido,
Estado_Civil = @Estado_Civil,
Direccion = @Direccion
where Num_Cedula = @Num_Cedula

create procedure Modificar_Detalle_Venta_Contado
@Id_Venta int,
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
if(exists(Select Id_Venta from Detalle_Venta_Contado where Id_Venta = @Id_Venta))
begin
update Detalle_Venta_Contado set Tipo = @Tipo,
Disponible = @Disponible,
Cantidad = @Cantidad,
Descuento = @Descuento,
Precio = dbo.Obtener_Precio(@Tipo,@Descuento),
SubTotal = dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),
Devolucion = dbo.Obtener_Devolucion(@Disponible,@Cantidad),
Comision_Vendedor = dbo.Calcular_Comision_Vendedor(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),
Comision_Arqueador = dbo.Calcular_Comision_Arqueador(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),
Comision_Administrador = dbo.Calcular_Comision_Administrador(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)))
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula and Id_Venta = @Id_Venta
update Venta_Contado set Total_C$ = (select sum(SubTotal) from Detalle_Venta_Contado where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula) 
where Fecha_Venta = Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Contado set Total_Comision_Vendedor = (select sum(Comision_Vendedor) from Detalle_Venta_Contado where Num_Cedula = @Num_Cedula and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Contado set Total_Comision_Arqueador = (select sum(Comision_Arqueador) from Detalle_Venta_Contado where Num_Cedula = @Num_Cedula and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Contado set Total_Comision_Administrador = (select sum(Comision_Administrador) from Detalle_Venta_Contado where Num_Cedula = @Num_Cedula and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula 
end
else
 Print 'El Id_Venta no existe'

-- Procedimientos de eliminación

create procedure Eliminar_Vendedor
@Num_Cedula char(16)
as
delete from Detalle_Venta_Contado where Num_Cedula = @Num_Cedula
delete from Venta_Contado where Num_Cedula = @Num_Cedula
delete from Vendedor where Num_Cedula = @Num_Cedula

create procedure Eliminar_Venta_Contado
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
delete from Venta_Contado where Fecha_Venta = @Fecha_Venta
 and Num_Cedula = @Num_Cedula

create procedure Eliminar_Detalle_Venta_Contado
@Id_Venta int,
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
if(exists(Select Id_Venta from Detalle_Venta_Contado where Id_Venta = @Id_Venta))
 begin
  update Venta_Contado set Total_Comision_Vendedor = Total_Comision_Vendedor - (select Comision_Vendedor from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
  where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
  
  update Venta_Contado set Total_Comision_Arqueador = Total_Comision_Arqueador - (select Comision_Arqueador from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
  where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
  
  update Venta_Contado set Total_Comision_Administrador = Total_Comision_Administrador - (select Comision_Administrador from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
  where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
  
  update Venta_Contado set Total_C$ = Total_C$ - (select SubTotal from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
  where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
  delete from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
 end
else
 Print 'Id_Venta no existe'

--------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------

create rule validar
as
@limite >= 0

sp_addlogin 'Agusto','nicaragua','MOVISTAR'
sp_addsrvrolemember 'Agusto',sysadmin

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

-- Procedimiento para generar un respaldo (backup)

create procedure respaldo
@bd as nvarchar(20),
@disp as nvarchar(20)
as
backup database @bd to @disp

-- Procedimiento para restaurar una BD

create procedure restaurar
@bd as nvarchar(20),
@disp as nvarchar(20)
as
restore database @bd from @disp

-- Procedimiento para Crear,Modificar y Eliminar un Usuario

create procedure Crear_Usuario
@Nombre nvarchar(35),
@Contraseña nvarchar(35),
@Tipo nvarchar(35)
as
insert into Usuario values(@Nombre,@Contraseña,@Tipo)

create procedure Eliminar_Usuario
@Id_Usuario int
as
delete from Usuario where Id_Usuario = @Id_Usuario

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

sp_addlogin 'Cesar','nicaragua','CLARO'
sp_addsrvrolemember 'Cesar',sysadmin

sp_addlogin 'Mayling','isabel','CLARO'
sp_addsrvrolemember 'Mayling',sysadmin

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

-- Prueba

Ingresar_Vendedor '001-291186-0013Y','Masculino','José','Francisco','Rodríguez','Chávez','Soltero','Batahola Norte'

Ingresar_Venta_Contado '070318','001-291186-0013Y'

Ingresar_Venta_Contado '070319','001-291186-0013Y'

Ingresar_Detalle_Venta_Contado '20',10,5,5,'070319','001-291186-0013Y'

Modificar_Detalle_Venta_Contado 3,'30',10,4,5,'070319','001-291186-0013Y'

Eliminar_Detalle_Venta_Contado 51,'070319','001-291186-0013Y'

delete from Detalle_Venta_Contado
delete from Venta_Contado
delete from Vendedor

backup database MOVISTAR to disk = 'C:\MOVISTAR'

use master

-- ultimas modificaciones 07/04/2007

alter table Detalle_Venta_Contado alter column Tipo char(3) not null
