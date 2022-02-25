create database CLARO

restore database CLARO from disk = 'C:\CLARO'

use CLARO

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
Fecha_Venta smalldatetime not null,--primary key not null,
Total_$ money null,
Total_C$ money null,
Total_Comision_Vendedor money null,
Total_Comision_Arqueador money null,
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
Fecha_Venta smalldatetime not null, --foreign key references Venta_Contado(Fecha_Venta) not null,
Num_Cedula char(16) foreign key references Vendedor(Num_Cedula) not null
)

create table Venta_Credito
(
Id_Venta int identity(1,1) primary key not null,
Fecha_Venta smalldatetime not null,--primary key not null,
Total_$ money null,
Total_C$ money null,
Total_Comision_Vendedor money null,
Total_Comision_Arqueador money null,
Num_Cedula char(16) foreign key references Vendedor(Num_Cedula) not null
)

create table Detalle_Venta_Credito
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
Fecha_Venta smalldatetime not null,--foreign key references Venta_Credito(Fecha_Venta) not null,
Num_Cedula char(16) foreign key references Vendedor(Num_Cedula) not null,
Num_Factura char(5) check(Num_Factura like '[0-9][0-9][0-9][0-9][0-9]') not null
)

create table Estado_Moneda
(
Id_Estado int identity(1,1) primary key not null,
Tasa_Actual money not null
)

create table Cambio
(
Id_Cambio int identity(1,1) primary key not null,
Fecha_Cambio smalldatetime not null,
Cambio_Tasa money not null
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
 declare @precio as money
 declare @st as money
 set @precio = (select cast(@Tipo as money))
 set @st = (@precio * @Cantidad)
 return @st
end

create function Obtener_Precio(@Tipo char(3),@Descuento float)
returns money
begin
 declare @entero as money
 declare @porcentaje as float
 declare @valor as money
 declare @precio as money
 set @entero = (select cast(@Tipo as money)) -- 1.5
 set @porcentaje = @Descuento / 100 -- 0.05
 set @valor = @entero * @porcentaje -- 1.5 * 0.05
 set @precio = @entero - @valor -- 1.5 - 0.075
 return @precio -- 1.425
end

Ingresar_Venta_Contado '070223','001-291186-0013Y'

Ingresar_Detalle_Venta_Contado '1.5',4,2,5,'070223','001-291186-0013Y'

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

sp_bindrule 'validar','Venta_Contado.Total_$'
sp_bindrule 'validar','Venta_Contado.Total_C$'
sp_bindrule 'validar','Venta_Contado.Total_Comision_Vendedor'
sp_bindrule 'validar','Venta_Contado.Total_Comision_Arqueador'
sp_bindrule 'validar','Detalle_Venta_Contado.Disponible'
sp_bindrule 'validar','Detalle_Venta_Contado.Cantidad'
sp_bindrule 'validar','Detalle_Venta_Contado.Descuento'
sp_bindrule 'validar','Detalle_Venta_Contado.Devolucion'
sp_bindrule 'validar','Detalle_Venta_Contado.Comision_Vendedor'
sp_bindrule 'validar','Detalle_Venta_Contado.Comision_Arqueador'
sp_bindrule 'validar','Detalle_Venta_Contado.SubTotal'
sp_bindrule 'validar','Detalle_Venta_Contado.SubTotal_Cordobas'

create procedure Ingresar_Detalle_Venta_Contado
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
insert into Detalle_Venta_Contado values(@Tipo,@Disponible,@Cantidad,@Descuento,dbo.Obtener_Precio(@Tipo,@Descuento),dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),dbo.Obtener_Devolucion(@Disponible,@Cantidad),dbo.Calcular_Comision_Vendedor(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),dbo.Calcular_Comision_Arqueador(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),@Fecha_Venta,@Num_Cedula,dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)))

create function Obtener_SubTotal_Cordobas(@Cantidad int,@Precio money)
returns money
begin
 declare @st as money
 declare @cord as money
 set @st = (@Cantidad * @Precio)
 set @cord = @st * (Select Tasa_Actual from Estado_Moneda where Id_Estado = 1)
 return @cord
end

create function Calcular_Comision_Vendedor(@Descuento float,@SubTotal money)
returns money
begin
 declare @sal as money
 if(@Descuento = 4)--5
  set @sal = (@SubTotal * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1)) * 0.0275
 if(@Descuento = 5)--6
  set @sal = (@SubTotal * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1)) * 0.0175
 return @sal
end

create function Calcular_Comision_Arqueador(@Descuento float,@SubTotal money)
returns money
begin
 declare @sal as money
 set @sal = (@SubTotal * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1)) * 0.0025
 return @sal
end

create trigger Actualizar_Venta_Contado
on Detalle_Venta_Contado after insert
as
update Venta_Contado set Total_$ = Total_$ + (select SubTotal from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Contado set Total_C$ = Total_C$ + (select SubTotal_Cordobas from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Contado set Total_Comision_Vendedor = Total_Comision_Vendedor + (select Comision_Vendedor from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Contado set Total_Comision_Arqueador = Total_Comision_Arqueador + (select Comision_Arqueador from inserted)
from inserted i,Venta_Contado v
where v.Fecha_Venta = i.Fecha_Venta and v.Num_Cedula = i.Num_Cedula

create procedure Ingresar_Venta_Credito
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
insert into Venta_Credito values(@Fecha_Venta,0,0,0,0,@Num_Cedula)

create procedure Ingresar_Detalle_Venta_Credito
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime,
@Num_Cedula char(16),
@Num_Factura char(5)
as
insert into Detalle_Venta_Credito values(@Tipo,@Disponible,@Cantidad,@Descuento,dbo.Obtener_Precio(@Tipo,@Descuento),dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),dbo.Obtener_Devolucion(@Disponible,@Cantidad),dbo.Calcular_Comision_Vendedor_Credito(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),dbo.Calcular_Comision_Arqueador(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),@Fecha_Venta,@Num_Cedula,@Num_Factura,dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)))

sp_bindrule 'validar','Venta_Credito.Total_$'
sp_bindrule 'validar','Venta_Credito.Total_C$'
sp_bindrule 'validar','Venta_Credito.Total_Comision_Vendedor'
sp_bindrule 'validar','Venta_Credito.Total_Comision_Arqueador'
sp_bindrule 'validar','Detalle_Venta_Credito.Disponible'
sp_bindrule 'validar','Detalle_Venta_Credito.Cantidad'
sp_bindrule 'validar','Detalle_Venta_Credito.Descuento'
sp_bindrule 'validar','Detalle_Venta_Credito.Devolucion'
sp_bindrule 'validar','Detalle_Venta_Credito.Comision_Vendedor'
sp_bindrule 'validar','Detalle_Venta_Credito.Comision_Arqueador'
sp_bindrule 'validar','Detalle_Venta_Credito.SubTotal'
sp_bindrule 'validar','Detalle_Venta_Credito.SubTotal_Cordobas'

create trigger Actualizar_Venta_Credito
on Detalle_Venta_Credito after insert
as
update Factura_Pendiente set Saldo_Pendiente = Saldo_Pendiente + (select SubTotal_Cordobas from inserted)
where Num_Cedula = (select Num_Cedula from inserted) and Num_Factura = (select Num_Factura from inserted)
update Factura_Pendiente set Comision_Vendedor = Comision_Vendedor + (select Comision_Vendedor from inserted)
where Num_Cedula = (select Num_Cedula from inserted) and Num_Factura = (select Num_Factura from inserted)
update Factura_Pendiente set Comision_Arqueador = Comision_Arqueador + (select Comision_Arqueador from inserted)
where Num_Cedula = (select Num_Cedula from inserted) and Num_Factura = (select Num_Factura from inserted)
update Venta_Credito set Total_$ = Total_$ + (select SubTotal from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Credito set Total_C$ = Total_C$ + (select SubTotal_Cordobas from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Credito set Total_Comision_Vendedor = Total_Comision_Vendedor + (select Comision_Vendedor from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted) and Num_Cedula = (select Num_Cedula from inserted)
update Venta_Credito set Total_Comision_Arqueador = Total_Comision_Arqueador + (select Comision_Arqueador from inserted)
from inserted i,Venta_Credito v
where v.Fecha_Venta = i.Fecha_Venta  and v.Num_Cedula = i.Num_Cedula

create procedure Ingresar_Cambio
@Fecha_Cambio smalldatetime,
@Cambio_Tasa money
as
insert into Cambio values(@Fecha_Cambio,@Cambio_Tasa)

sp_bindrule 'validar','Estado_Moneda.Tasa_Actual'
sp_bindrule 'validar','Cambio.Cambio_Tasa'

create trigger Actualizar_Estado_Moneda
on Cambio after insert
as
update Estado_Moneda set Tasa_Actual = (select Cambio_Tasa from inserted)
from inserted i, Estado_Moneda e
where e.Id_Estado = 1

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
if(exists(select Id_Venta from Detalle_Venta_Contado where Id_Venta = @Id_Venta))
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
SubTotal_Cordobas = dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula and Id_Venta = @Id_Venta
update Venta_Contado set Total_$ = (select sum(SubTotal) from Detalle_Venta_Contado where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Contado set Total_C$ = (select sum(SubTotal_Cordobas) from Detalle_Venta_Contado where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Contado set Total_Comision_Vendedor = (select sum(Comision_Vendedor) from Detalle_Venta_Contado where Num_Cedula = @Num_Cedula and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Contado set Total_Comision_Arqueador = (select sum(Comision_Arqueador) from Detalle_Venta_Contado where Num_Cedula = @Num_Cedula and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
end
else
 Print 'El Id_Venta no existe'

create procedure Modificar_Detalle_Venta_Credito
@Id_Venta int,
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime,
@Num_Cedula char(16),
@Num_Factura char(5)
as
if(exists(select Id_Venta from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Num_Factura = @Num_Factura))
begin
update Detalle_Venta_Credito set Tipo = @Tipo,
Disponible = @Disponible,
Cantidad = @Cantidad,
Descuento = @Descuento,
Precio = dbo.Obtener_Precio(@Tipo,@Descuento),
SubTotal = dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),
Devolucion = dbo.Obtener_Devolucion(@Disponible,@Cantidad),
Comision_Vendedor = dbo.Calcular_Comision_Vendedor_Credito(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),
Comision_Arqueador = dbo.Calcular_Comision_Arqueador(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),
SubTotal_Cordobas = dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula and Id_Venta = @Id_Venta
update Venta_Credito set Total_$ = (select sum(SubTotal) from Detalle_Venta_Credito where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Credito set Total_C$ = (select sum(SubTotal_Cordobas) from Detalle_Venta_Credito where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Credito set Total_Comision_Arqueador = (select sum(Comision_Arqueador) from Detalle_Venta_Credito where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
update Venta_Credito set Total_Comision_Vendedor = (select sum(Comision_Vendedor) from Detalle_Venta_Credito where Num_Cedula = @Num_Cedula and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
update Factura_Pendiente set Saldo_Pendiente = (select sum(SubTotal_Cordobas) from Detalle_Venta_Credito where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura)
where Fecha_Factura = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura
update Factura_Pendiente set Comision_Arqueador = (select sum(Comision_Arqueador) from Detalle_Venta_Credito where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura)
where Fecha_Factura = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura
update Factura_Pendiente set Comision_Vendedor = (select sum(Comision_Vendedor) from Detalle_Venta_Credito where Num_Cedula = @Num_Cedula and Fecha_Venta = @Fecha_Venta and Num_Factura = @Num_Factura)
where Fecha_Factura = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura
end
else
 Print 'El Id_Venta no existe'

create procedure Modificar_Cambio
@Fecha_Cambio smalldatetime,
@Cambio_Tasa money
as
update Cambio set Cambio_Tasa = Cambio_Tasa
where Fecha_Cambio = @Fecha_Cambio
update Estado_Moneda set Tasa_Actual = (select Cambio_Tasa from Cambio where Fecha_Cambio = @Fecha_Cambio)
where Id_Estado = 1

create procedure Modificar_Abono
@Fecha_Abono smalldatetime,
@Cantidad_C$ money,
@Num_Cedula char(16)
as
update Abono set Cantidad_C$ = @Cantidad_C$ 
where Fecha_Abono = @Fecha_Abono and Num_Cedula = @Num_Cedula

-- Procedimientos de eliminación

create procedure Eliminar_Vendedor
@Num_Cedula char(16)
as
delete from Factura_Pendiente where Num_Cedula = @Num_Cedula
delete from Detalle_Venta_Contado where Num_Cedula = @Num_Cedula
delete from Detalle_Venta_Credito where Num_Cedula = @Num_Cedula
delete from Venta_Contado where Num_Cedula = @Num_Cedula
delete from Venta_Credito where Num_Cedula = @Num_Cedula
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
if(exists(select Id_Venta from Detalle_Venta_Contado where Id_Venta = @Id_Venta))
begin
 update Venta_Contado set Total_Comision_Vendedor = Total_Comision_Vendedor - (select Comision_Vendedor from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
 where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
 update Venta_Contado set Total_Comision_Arqueador = Total_Comision_Arqueador - (select Comision_Arqueador from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
 where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
 update Venta_Contado set Total_$ = Total_$ - (select SubTotal from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
 where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
 update Venta_Contado set Total_C$ = Total_C$ - (select SubTotal_Cordobas from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
 where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
 delete from Detalle_Venta_Contado where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula
end
else
 Print 'El Id_Venta no existe'

if(exists(Select Id_Venta from Detalle_Venta_Contado where Id_Venta = 511))
 Print 'Existe'
else
 Print 'No Existe'

Eliminar_Detalle_Venta_Contado 1,'070522','001-291186-0013Y'

create procedure Eliminar_Venta_Credito
@Fecha_Venta smalldatetime,
@Num_Cedula char(16)
as
delete from Venta_Credito 
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula

create procedure Eliminar_Detalle_Venta_Credito
@Id_Venta int,
@Fecha_Venta smalldatetime,
@Num_Cedula char(16),
@Num_Factura char(5)
as
if(exists(select Id_Venta from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Num_Factura = @Num_Factura))
begin
update Venta_Credito set Total_Comision_Vendedor = Total_Comision_Vendedor - (select Comision_Vendedor from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula  
update Venta_Credito set Total_Comision_Arqueador = Total_Comision_Arqueador - (select Comision_Arqueador from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula  
update Venta_Credito set Total_$ = Total_$ - (select SubTotal from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula  
update Venta_Credito set Total_C$ = Total_C$ - (select SubTotal_Cordobas from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula  
update Factura_Pendiente set Saldo_Pendiente = Saldo_Pendiente - (select SubTotal_Cordobas from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Factura = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura  
update Factura_Pendiente set Comision_Vendedor = Comision_Vendedor - (select Comision_Vendedor from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Factura = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura  
update Factura_Pendiente set Comision_Arqueador = Comision_Arqueador - (select Comision_Arqueador from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula)
where Fecha_Factura = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura  
delete from Detalle_Venta_Credito where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta and Num_Cedula = @Num_Cedula and Num_Factura = @Num_Factura
end
else
 Print 'El Id_Venta no existe'

create procedure Eliminar_Cambio
@Fecha_Cambio smalldatetime
as
delete from Cambio where Fecha_Cambio = @Fecha_Cambio

--------------------------------------------------------------------------------------
--------------------------------------------------------------------------------------

create rule validar
as
@limite >= 0

sp_addlogin 'Cesar','Agusto','CLARO'
sp_addsrvrolemember 'Cesar',sysadmin

delete from Venta_Contado
delete from Detalle_Venta_Contado

delete from Venta_Credito
delete from Detalle_Venta_Credito

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

select * from Detalle_Venta_Contado where Fecha_Venta >= '070215' and Fecha_Venta <= '070217' and Num_Cedula = '001-291186-0013Y'

delete from Detalle_Venta_Contado
delete from Detalle_Venta_Credito

delete from Venta_Contado
delete from Venta_Credito

delete from Vendedor

-- Prueba

--------------------------------------------

Ingresar_Vendedor '001-291186-0013Y','Masculino','José','Francisco','Rodríguez','Chávez','Soltero','Batahola Norte 808'

Ingresar_Credito '001-291186-0013Y'

Ingresar_Salario_Vendedor '001-291186-0013Y'

--------------------------------------------

Ingresar_Venta_Contado '070222','001-291186-0013Y'

Ingresar_Venta_Credito '070222','001-291186-0013Y'

Ingresar_Detalle_Venta_Contado '1.5',5,3,5,'070222','001-291186-0013Y'

Ingresar_Detalle_Venta_Credito '1',5,3,5,'070222','001-291186-0013Y','12345'

Ingresar_Detalle_Salario_Vendedor '070222',1,'001-291186-0013Y'

Modificar_Detalle_Salario_Vendedor 22,'070222',0,'001-291186-0013Y'

Eliminar_Detalle_Salario_Vendedor 22,'070222','001-291186-0013Y'

Eliminar_Detalle_Venta_Contado 39,'070222','001-291186-0013Y'

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

Ingresar_Detalle_Salario_Arqueador '070222',0.0001,'1234'

Modificar_Detalle_Salario_Arqueador 23,'070222',0.0002,'1234'

Eliminar_Detalle_Salario_Arqueador 23,'070222','1234'

Ingresar_Detalle_Credito '070222',1,'001-291186-0013Y'

Modificar_Detalle_Credito 11,'070222',2,'001-291186-0013Y'

Eliminar_Detalle_Credito 11,'070222','001-291186-0013Y'

create table Venta_Contado_Admon
(
Fecha_Venta smalldatetime primary key not null,
Total_$ money null,
Total_C$ money null,
Total_Comision_Admon money null
)

create table Detalle_Venta_Contado_Admon
(
Id_Venta int identity(1,1) primary key not null,
Tipo char(3) not null,
Disponible int not null,
Cantidad int not null,
Descuento float not null,
Precio money null,
SubTotal money null,
Devolucion int null,
Comision_Admon money null,
Fecha_Venta smalldatetime foreign key references Venta_Contado_Admon(Fecha_Venta) not null
)

create table Venta_Credito_Admon
(
Fecha_Venta smalldatetime primary key not null,
Total_$ money null,
Total_C$ money null,
Total_Comision_Admon money null
)

create table Detalle_Venta_Credito_Admon
(
Id_Venta int identity(1,1) primary key not null,
Tipo char(3) not null,
Disponible int not null,
Cantidad int not null,
Descuento float not null,
Precio money null,
SubTotal money null,
Devolucion int null,
Comision_Admon money null,
Fecha_Venta smalldatetime foreign key references Venta_Credito_Admon(Fecha_Venta) not null,
Num_Factura char(5) check(Num_Factura like '[0-9][0-9][0-9][0-9][0-9]') not null,
)

create procedure Ingresar_Venta_Contado_Admon
@Fecha_Venta smalldatetime
as
insert into Venta_Contado_Admon values(@Fecha_Venta,0,0,0)

sp_bindrule 'validar','Venta_Contado_Admon.Total_$'
sp_bindrule 'validar','Venta_Contado_Admon.Total_C$'
sp_bindrule 'validar','Venta_Contado_Admon.Total_Comision_Admon'
sp_bindrule 'validar','Detalle_Venta_Contado_Admon.Disponible'
sp_bindrule 'validar','Detalle_Venta_Contado_Admon.Cantidad'
sp_bindrule 'validar','Detalle_Venta_Contado_Admon.Descuento'
sp_bindrule 'validar','Detalle_Venta_Contado_Admon.Devolucion'
sp_bindrule 'validar','Detalle_Venta_Contado_Admon.Comision_Admon'
sp_bindrule 'validar','Detalle_Venta_Contado.SubTotal'

create procedure Ingresar_Detalle_Venta_Contado_Admon
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime
as
insert into Detalle_Venta_Contado_Admon values(@Tipo,@Disponible,@Cantidad,@Descuento,dbo.Obtener_Precio(@Tipo,@Descuento),dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),dbo.Obtener_Devolucion(@Disponible,@Cantidad),dbo.Calcular_Comision_Admon(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),@Fecha_Venta,dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)))

create function Calcular_Comision_Admon(@Descuento float,@SubTotal money)
returns money
begin
 declare @sal as money
 if(@Descuento = 5)
  set @sal = (@SubTotal * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1)) * 0.03
 if(@Descuento = 6)
  set @sal = (@SubTotal * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1)) * 0.02
 return @sal
end

create trigger Actualizar_Venta_Contado_Admon
on Detalle_Venta_Contado_Admon after insert
as
update Venta_Contado_Admon set Total_$ = Total_$ + (select SubTotal from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted)
update Venta_Contado_Admon set Total_C$ = Total_C$ + (select SubTotal_Cordobas from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted)
update Venta_Contado_Admon set Total_Comision_Admon = Total_Comision_Admon + (select Comision_Admon from inserted) 
from inserted i,Venta_Contado_Admon v
where v.Fecha_Venta = i.Fecha_Venta

create procedure Ingresar_Venta_Credito_Admon
@Fecha_Venta smalldatetime
as
insert into Venta_Credito_Admon values(@Fecha_Venta,0,0,0)

sp_bindrule 'validar','Venta_Credito_Admon.Total_$'
sp_bindrule 'validar','Venta_Credito_Admon.Total_C$'
sp_bindrule 'validar','Detalle_Venta_Credito_Admon.Disponible'
sp_bindrule 'validar','Detalle_Venta_Credito_Admon.Cantidad'
sp_bindrule 'validar','Detalle_Venta_Credito_Admon.Descuento'
sp_bindrule 'validar','Detalle_Venta_Credito_Admon.Devolucion'
sp_bindrule 'validar','Detalle_Venta_Credito_Admon.Comision_Admon'
sp_bindrule 'validar','Detalle_Venta_Credito.SubTotal'

create procedure Ingresar_Detalle_Venta_Credito_Admon
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime,
@Num_Factura char(5)
as
insert into Detalle_Venta_Credito_Admon values(@Tipo,@Disponible,@Cantidad,@Descuento,dbo.Obtener_Precio(@Tipo,@Descuento),dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),dbo.Obtener_Devolucion(@Disponible,@Cantidad),dbo.Calcular_Comision_Admon(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),@Fecha_Venta,@Num_Factura,dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)))

create trigger Actualizar_Venta_Credito_Admon
on Detalle_Venta_Credito_Admon after insert
as
update Factura_Pendiente_Admon set Saldo_Pendiente = Saldo_Pendiente + (select SubTotal_Cordobas from inserted)
where Fecha_Factura = (select Fecha_Venta from inserted) and Num_Factura = (select Num_Factura from inserted)
update Factura_Pendiente_Admon set Comision_Admon = Comision_Admon + (select Comision_Admon from inserted)
where Fecha_Factura = (select Fecha_Venta from inserted) and Num_Factura = (select Num_Factura from inserted)
update Venta_Credito_Admon set Total_$ = Total_$ + (select SubTotal from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted)
update Venta_Credito_Admon set Total_C$ = Total_C$ + (select SubTotal_Cordobas from inserted)
where Fecha_Venta = (select Fecha_Venta from inserted)
update Venta_Credito_Admon set Total_Comision_Admon = Total_Comision_Admon + (select Comision_Admon from inserted)
from inserted i,Venta_Credito_Admon v
where v.Fecha_Venta = i.Fecha_Venta

backup database CLARO to disk = 'C:\CLARO'

create procedure Modificar_Detalle_Venta_Contado_Admon
@Id_Venta int,
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime
as
if(exists(select Id_Venta from Detalle_Venta_Contado_Admon where Id_Venta = @Id_Venta))
begin
update Detalle_Venta_Contado_Admon set Tipo = @Tipo,
Disponible = @Disponible,
Cantidad = @Cantidad,
Descuento = @Descuento,
Precio = dbo.Obtener_Precio(@Tipo,@Descuento),
SubTotal = dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),
Devolucion = dbo.Obtener_Devolucion(@Disponible,@Cantidad),
Comision_Admon = dbo.Calcular_Comision_Admon(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),
SubTotal_Cordobas = dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))
where Fecha_Venta = @Fecha_Venta and Id_Venta = @Id_Venta
update Venta_Contado_Admon set Total_$ = (select sum(SubTotal) from Detalle_Venta_Contado_Admon where Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Contado_Admon set Total_C$ = (select sum(SubTotal_Cordobas) from Detalle_Venta_Contado_Admon where Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Contado_Admon set Total_Comision_Admon = (select sum(Comision_Admon) from Detalle_Venta_Contado_Admon where Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
end
else
 Print 'El Id_Venta no existe'

create procedure Modificar_Detalle_Venta_Credito_Admon
@Id_Venta int,
@Tipo char(3),
@Disponible int,
@Cantidad int,
@Descuento float,
@Fecha_Venta smalldatetime,
@Num_Factura char(5)
as
if(exists(select Id_Venta from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta))
begin
update Detalle_Venta_Credito_Admon set Tipo = @Tipo,
Disponible = @Disponible,
Cantidad = @Cantidad,
Descuento = @Descuento,
Precio = dbo.Obtener_Precio(@Tipo,@Descuento),
SubTotal = dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento)),
Devolucion = dbo.Obtener_Devolucion(@Disponible,@Cantidad),
Comision_Admon = dbo.Calcular_Comision_Admon(@Descuento,dbo.Obtener_SubTotal(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))),
SubTotal_Cordobas = dbo.Obtener_SubTotal_Cordobas(@Cantidad,dbo.Obtener_Precio(@Tipo,@Descuento))
where Fecha_Venta = @Fecha_Venta and Id_Venta = @Id_Venta
update Venta_Credito_Admon set Total_$ = (select sum(SubTotal) from Detalle_Venta_Credito_Admon where Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Credito_Admon set Total_C$ = (select sum(SubTotal_Cordobas) from Detalle_Venta_Credito_Admon where Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Credito_Admon set Total_Comision_Admon = (select sum(Comision_Admon) from Detalle_Venta_Credito_Admon where Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Factura_Pendiente_Admon set Saldo_Pendiente = (select sum(SubTotal_Cordobas) from Detalle_Venta_Credito_Admon where Fecha_Venta = @Fecha_Venta and Num_Factura = @Num_Factura)
where Fecha_Factura = @Fecha_Venta and Num_Factura = @Num_Factura
update Factura_Pendiente_Admon set Comision_Admon = (select sum(Comision_Admon) from Detalle_Venta_Credito_Admon where Fecha_Venta = @Fecha_Venta and Num_Factura = @Num_Factura)
where Fecha_Factura = @Fecha_Venta and Num_Factura = @Num_Factura
end
else
 Print 'El Id_Venta no existe'

create procedure Eliminar_Venta_Contado_Admon
@Fecha_Venta smalldatetime
as
delete from Venta_Contado_Admon where Fecha_Venta = @Fecha_Venta

create procedure Eliminar_Detalle_Venta_Contado_Admon
@Id_Venta int,
@Fecha_Venta smalldatetime
as
update Venta_Contado_Admon set Total_Comision_Admon = Total_Comision_Admon - (select Comision_Admon from Detalle_Venta_Contado_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Contado_Admon set Total_$ = Total_$ - (select SubTotal from Detalle_Venta_Contado_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Contado_Admon set Total_C$ = Total_C$ - ((select SubTotal from Detalle_Venta_Contado_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta) * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1))
where Fecha_Venta = @Fecha_Venta
delete from Detalle_Venta_Contado_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta

create procedure Eliminar_Venta_Credito_Admon
@Fecha_Venta smalldatetime
as
delete from Venta_Credito_Admon 
where Fecha_Venta = @Fecha_Venta

create procedure Eliminar_Detalle_Venta_Credito_Admon
@Id_Venta int,
@Fecha_Venta smalldatetime,
@Num_Factura char(5)
as
if(exists(select Id_Venta from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta))
begin
update Venta_Credito_Admon set Total_Comision_Admon = Total_Comision_Admon - (select Comision_Admon from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Credito_Admon set Total_$ = Total_$ - (select SubTotal from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta)
where Fecha_Venta = @Fecha_Venta
update Venta_Credito_Admon set Total_C$ = Total_C$ - ((select SubTotal from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta) * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1))
where Fecha_Venta = @Fecha_Venta
update Factura_Pendiente_Admon set Saldo_Pendiente = Saldo_Pendiente - ((select SubTotal from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta) * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1))
where Fecha_Factura = @Fecha_Venta and Num_Factura = @Num_Factura
update Factura_Pendiente_Admon set Comision_Admon = Comision_Admon - (select Comision_Admon from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta)
where Fecha_Factura = @Fecha_Venta and Num_Factura = @Num_Factura
delete from Detalle_Venta_Credito_Admon where Id_Venta = @Id_Venta and Fecha_Venta = @Fecha_Venta
end
else
 Print 'El Id_Venta no existe'

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

create procedure Modificar_Contraseña
@Nombre nvarchar(35),
@Contraseña nvarchar(35)
as
update Usuario set Contraseña = @Contraseña where Nombre = @Nombre

delete from Detalle_Venta_Contado_Admon
delete from Detalle_Venta_Credito_Admon

delete from Venta_Contado_Admon
delete from Venta_Credito_Admon

Ingresar_Venta_Credito_Admon '070222'

Ingresar_Detalle_Venta_Credito_Admon '1',50,25,5,'070222' -- 12.8545 + 7.7127 = 20.5672

Modificar_Detalle_Venta_Credito_Admon 20,'1',50,15,5,'070222' -- 25.709

Eliminar_Detalle_Venta_Credito_Admon 39,'070222'

delete from Cambio

delete from Detalle_Credito

delete from Vendedor

backup database CLARO to disk = 'C:\CLARO'

create procedure Eliminar_Cambio
@Id_Cambio int
as
delete from Cambio where Id_Cambio = @Id_Cambio

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

-- Últimas modificaciones Sábado, 03 de Febrero del 2007

delete from Detalle_Venta_Contado
delete from Venta_Contado
delete from Detalle_Venta_Credito
delete from Venta_Credito
delete from Vendedor

delete from Detalle_Salario_Arqueador

-- ILICH

Ingresar_Vendedor '001-291186-0013Y','Masculino','Ilich','Francisco','Rodríguez','Chávez','Casado','Batahola Norte 808'

Ingresar_Credito '001-291186-0013Y'

Ingresar_Venta_Credito '070303','001-291186-0013Y'

Ingresar_Detalle_Venta_Credito '1',2,1,5,'070303','001-291186-0013Y'

Modificar_Detalle_Venta_Credito 36,'1',2,2,5,'070303','001-291186-0013Y'

Venta C$: 17.2710
Salario C$: 0.4750
Arqueo C$: 0.0432
Deuda C$: 16.7528

Ingresar_Venta_Contado '070303','001-291186-0013Y'

Ingresar_Detalle_Venta_Contado '1',2,1,5,'070303','001-291186-0013Y'

Modificar_Detalle_Venta_Contado 10,'1',2,2,5,'070303','001-291186-0013Y'

Eliminar_Detalle_Venta_Contado 10,'070303','001-291186-0013Y'

-- CESAR

Ingresar_Vendedor '001-291086-0013Y','Masculino','Cesar','Francisco','Rodríguez','Chávez','Casado','Batahola Norte 808'

Ingresar_Salario_Vendedor '001-291086-0013Y'

Ingresar_Credito '001-291086-0013Y'

Ingresar_Venta_Credito '070303','001-291086-0013Y'

Ingresar_Detalle_Venta_Credito '1',2,1,5,'070303','001-291086-0013Y'


Ingresar_Venta_Contado '070303','001-291086-0013Y'

Ingresar_Detalle_Venta_Contado '1',2,1,5,'070303','001-291086-0013Y'

Modificar_Detalle_Venta_Contado 12,'1',2,2,5,'070303','001-291086-0013Y'

Eliminar_Detalle_Venta_Contado 12,'070303','001-291086-0013Y'

Venta C$: 17.1393
Salario C$: 0.4713
Arqueo C$: 0.0428
Deuda C$: 16.6252

-- JAVIER

Ingresar_Vendedor '001-291286-0013Y','Masculino','Javier','Francisco','Rodríguez','Chávez','Casado','Batahola Norte 808'

Ingresar_Salario_Vendedor '001-291286-0013Y'

Ingresar_Credito '001-291286-0013Y'

Ingresar_Venta_Contado '070303','001-291286-0013Y'

Ingresar_Detalle_Venta_Contado '1',2,1,5,'070303','001-291286-0013Y'

Modificar_Detalle_Venta_Contado 8,'1',2,2,5,'070303','001-291286-0013Y'

Venta C$: 17.1393
Salario C$: 0.4713
Arqueo C$: 0.0428
Deuda C$: 16.6252

-- MAYLING

Salario Total C$: 0.1284
0.1712

Modificar_Detalle_Venta_Credito 18,'1',2,1,5,'070303','001-291286-0013Y'
Eliminar_Detalle_Venta_Credito 18,'070303','001-291286-0013Y'

Ingresar_Detalle_Salario_Arqueador '070303',0.0856,'1234'

Ingresar_Detalle_Salario_Vendedor '070303',0.0013,'001-291286-0013Y'

Ingresar_Detalle_Credito '070303',0.0013,'001-291286-0013Y'

Modificar_Detalle_Credito 16,'070303',0,'001-291286-0013Y'

Eliminar_Detalle_Credito 17,'070303','001-291286-0013Y'

-- Generar Respaldo

backup database CLARO to disk = 'C:\CLARO'

Ingresar_Vendedor '401-060984-0007W','Masculino','José','Javier','Martínez','Perez','Casado','Masaya'

Ingresar_Salario_Vendedor '401-060984-0007W'

Ingresar_Credito '401-060984-0007W'

restore database CLARO from disk = 'C:\CLARO'

use CLARO

select * from Venta_Contado where Num_Cedula = '001-291186-0013Y' And Fecha_Venta = '070301'

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

sp_addlogin 'Cesar','nicaragua','CLARO'
sp_addsrvrolemember 'Cesar',sysadmin

sp_addlogin 'Mayling','isabel','CLARO'
sp_addsrvrolemember 'Mayling',sysadmin

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

create table Factura_Pendiente
(
Num_Factura char(5) check(Num_Factura like '[0-9][0-9][0-9][0-9][0-9]') primary key not null,
Fecha_Factura smalldatetime not null,
Saldo_Pendiente money null,
Comision_Vendedor money null,
Comision_Arqueador money null,
Estado_Factura nvarchar(25) null,
Num_Cedula char(16) foreign key references Vendedor(Num_Cedula) not null
)

sp_bindrule 'validar','Factura_Pendiente.Saldo_Pendiente'
sp_bindrule 'validar','Factura_Pendiente.Comision_Vendedor'
sp_bindrule 'validar','Factura_Pendiente.Comision_Arqueador'

create procedure Ingresar_Factura_Pendiente
@Num_Factura char(5),
@Fecha_Factura smalldatetime,
@Num_Cedula char(16)
as
insert into Factura_Pendiente values(@Num_Factura,@Fecha_Factura,0,0,0,'Pendiente',@Num_Cedula)

create procedure Modificar_Factura_Pendiente
@Num_Factura char(5),
@Fecha_Factura smalldatetime
as
if(exists(select Num_Factura from Factura_Pendiente where Num_Factura = @Num_Factura))
 begin
  update Factura_Pendiente set Estado_Factura = 'Cancelado',
  Fecha_Factura = @Fecha_Factura
  where Num_factura = @Num_Factura
 end
else
 Print 'El Num_Factura no existe'

create procedure Eliminar_Factura_Pendiente
@Num_Factura char(4)
as
if(exists(Select Num_Factura from Factura_Pendiente))
 begin
  delete from Factura_Pendiente where Num_Factura = @Num_Factura
 end
else
 Print 'El Num_Factura no existe'

create function Calcular_Comision_Vendedor_Credito(@Descuento float,@SubTotal money)
returns money
begin
 declare @sal as money
  set @sal = (@SubTotal * (select Tasa_Actual from Estado_Moneda where Id_Estado = 1)) * 0.0175
 return @sal
end

---------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------

create table Factura_Pendiente_Admon
(
Num_Factura char(5) check(Num_Factura like '[0-9][0-9][0-9][0-9][0-9]') primary key not null,
Fecha_Factura smalldatetime not null,
Saldo_Pendiente money null,
Comision_Admon money null,
Estado_Factura nvarchar(25) null
)

sp_bindrule 'validar','Factura_Pendiente_Admon.Saldo_Pendiente'
sp_bindrule 'validar','Factura_Pendiente_Admon.Comision_Admon'

create procedure Ingresar_Factura_Pendiente_Admon
@Num_Factura char(5),
@Fecha_Factura smalldatetime
as
insert into Factura_Pendiente_Admon values(@Num_Factura,@Fecha_Factura,0,0,'Pendiente')

create procedure Modificar_Factura_Pendiente_Admon
@Num_Factura char(5),
@Fecha_Factura smalldatetime
as
if(exists(select Num_Factura from Factura_Pendiente_Admon where Num_Factura = @Num_Factura))
 begin
  update Factura_Pendiente_Admon set Estado_Factura = 'Cancelado',
  Fecha_Factura = @Fecha_Factura
  where Num_Factura = @Num_Factura
 end
else
 Print 'El Num_Factura no existe'

create procedure Eliminar_Factura_Pendiente_Admon
@Num_Factura char(4)
as
if(exists(Select Num_Factura from Factura_Pendiente_Admon))
 begin
  delete from Factura_Pendiente_Admon where Num_Factura = @Num_Factura
 end
else
 Print 'El Num_Factura no existe'

delete from Detalle_Venta_Contado
delete from Venta_Contado

delete from Detalle_Venta_Credito
delete from Venta_Credito

delete from Detalle_Venta_Contado_Admon
delete from Venta_Contado_Admon

delete from Detalle_Venta_Credito_Admon
delete from Venta_Credito_Admon

delete from Factura_Pendiente
delete from Factura_Pendiente_Admon

Ingresar_Detalle_Venta_Credito '1',10,8,5,'070317','001-120484-0012N'

Ingresar_Venta_Contado_Admon '070317'

Ingresar_Venta_Contado_Admon '070320'

Ingresar_Detalle_Venta_Contado_Admon '1',132,132,5,'070320'

Eliminar_Detalle_Venta_Contado_Admon 31,'070320'



Ingresar_Venta_Credito_Admon '070317'

Ingresar_Venta_Credito_Admon '070320'

Ingresar_Detalle_Venta_Credito_Admon '1',10,5,5,'070320'

Eliminar_Detalle_Venta_Credito_Admon 11,'070320'

delete from Factura_Cancelada_Admon

backup database CLARO to disk = 'C:\CLARO'

use master

select * from Factura_Pendiente where Estado_Factura = 'Pendiente'

alter table Detalle_Venta_Contado alter column Tipo char(3) not null
alter table Detalle_Venta_Credito alter column Tipo char(3) not null

alter table Detalle_Venta_Contado_Admon alter column Tipo char(3) not null
alter table Detalle_Venta_Credito_Admon alter column Tipo char(3) not null

delete from Venta_Contado
delete from Detalle_Venta_Contado

use master

use CLARO

backup database CLARO to disk = 'C:\CLARO'

create database CLARO

restore database CLARO from disk = 'C:\CLARO'

delete from Factura_Pendiente
delete from Detalle_Venta_Credito
delete from Venta_Credito

-- Hoy 28/04/07 mi novia y yo cometimos una estupidez
-- USAR EL QUERY

alter table Detalle_Venta_Contado add SubTotal_Cordobas money

alter table Detalle_Venta_Contado alter column SubTotal_Cordobas money

update Detalle_Venta_Contado set SubTotal_Cordobas = SubTotal * 18.26 where Fecha_Venta between '070402' and '070415'

update Detalle_Venta_Contado set SubTotal_Cordobas = SubTotal * 18.30 where Fecha_Venta between '070416' and '070428'

<>

alter table Detalle_Venta_Credito add SubTotal_Cordobas money

alter table Detalle_Venta_Credito alter column SubTotal_Cordobas money

update Detalle_Venta_Credito set SubTotal_Cordobas = SubTotal * 18.26 where Fecha_Venta between '070402' and '070415'

update Detalle_Venta_Credito set SubTotal_Cordobas = SubTotal * 18.30 where Fecha_Venta between '070416' and '070428'

ILICH

update Venta_Credito set Total_$ = (Select sum(SubTotal) from Detalle_Venta_Credito where Num_Cedula = '401-040583-0005b')
where Num_Cedula = '401-040583-0005b'

update Venta_Credito set Total_C$ = (Select sum(SubTotal_Cordobas) from Detalle_Venta_Credito where Num_Cedula = '401-040583-0005b')
where Num_Cedula = '401-040583-0005b'

MAYCOL

update Venta_Credito set Total_$ = (Select sum(SubTotal) from Detalle_Venta_Credito where Num_Cedula = '401-050875-0004l')
where Num_Cedula = '401-050875-0004l'

update Venta_Credito set Total_C$ = (Select sum(SubTotal_Cordobas) from Detalle_Venta_Credito where Num_Cedula = '401-050875-0004l')
where Num_Cedula = '401-050875-0004l'

JOSE JAVIER

update Venta_Credito set Total_$ = (Select sum(SubTotal) from Detalle_Venta_Credito where Num_Cedula = '401-060984-0007W')
where Num_Cedula = '401-060984-0007W'

update Venta_Credito set Total_C$ = (Select sum(SubTotal_Cordobas) from Detalle_Venta_Credito where Num_Cedula = '401-060984-0007W')
where Num_Cedula = '401-060984-0007W'

JADER

update Venta_Credito set Total_$ = (Select sum(SubTotal) from Detalle_Venta_Credito where Num_Cedula = '401-160774-0005m')
where Num_Cedula = '401-160774-0005m'

update Venta_Credito set Total_C$ = (Select sum(SubTotal_Cordobas) from Detalle_Venta_Credito where Num_Cedula = '401-160774-0005m')
where Num_Cedula = '401-160774-0005m'

update Factura_Pendiente set Saldo_Pendiente = 
(Select sum(SubTotal_Cordobas) from Detalle_Venta_Credito 
where Num_Factura = '11003')
where Num_Factura = '11003'

backup database CLARO to disk = 'C:\CLARO'

delete from Detalle_Venta_Credito where Num_cedula = '401-160774-0005m'

update Factura_Pendiente set Saldo_Pendiente = (Select sum(SubTotal_Cordobas) from Detalle_Venta_Credito where Num_Factura = '11003')
where Num_Factura = '11003' 

sp_addlogin 'Mayling','isabel','CLARO'
sp_addsrvrolemember 'Mayling',sysadmin

alter table Detalle_Venta_Contado_Admon add SubTotal_Cordobas money null

alter table Detalle_Venta_Credito_Admon add SubTotal_Cordobas money null

sp_bindrule 'validar','Detalle_Venta_Contado_Admon.SubTotal_Cordobas'
sp_bindrule 'validar','Detalle_Venta_Credito_Admon.SubTotal_Cordobas'

select * from Factura_Pendiente_Admon where Fecha_Factura >= '070512' and Fecha_Factura <= '070513'

backup database CLARO to disk = 'C:\Respaldo 200507\CLARO'

backup database MOVISTAR to disk = 'C:\Respaldo 200507\MOVISTAR'

declare @fecha as smalldatetime
set @fecha = '070520'
select @fecha