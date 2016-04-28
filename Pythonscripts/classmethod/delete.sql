/*!40101 SET NAMES utf8 */; 
use dcf_user
go
delete from customer_reg_invitation where invited_customer_name='测试企业';
go
delete from user where customer_name='测试企业';
go
use dcf_customer
go
delete from customer where customer_name='测试企业';
go
delete from customer_base_info where customer_name='测试企业';



