# exceltosql
将excel转成sql

### excel格式
WZCKJLB  物资出库单							
序号	字段说明	字段名	类型	长度	默认值	允许为空	主键
1	    系统id	  id	   int	4			        √
2	申请单id	sqd_id	int	4			
3	申请单类型	sqtype	int	4			
4	验证码	yz_code	varchar	8			
5	操作人	czr	varchar	50			
6	操作时间	czsj	datetime				
7	预留字段6个

转成sql或直接执行

