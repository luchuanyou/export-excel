<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    <base href="<%=basePath%>">
    
    <title>用户信息</title>
    
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache">
	<meta http-equiv="expires" content="0">    
	<meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
	<meta http-equiv="description" content="This is my page">
	<!--
	<link rel="stylesheet" type="text/css" href="styles.css">
	-->

  </head>
  
  <body>
  	<div style="width: 200px;margin: auto;">
  		<table border="1px"  bordercolor="blue" cellpadding="0" cellspacing="0">
  			<tbody>用户信息</tbody>
  			<tr>
  				<td>错误码：</td>
  				<td>${errorCode }</td>
  			</tr>
  			<tr>
  				<td>错误描述：</td>
  				<td>${errorMsg }</td>
  			</tr>
  			<tr>
  				<td>用户名：</td>
  				<td>${memberInfoDTO.registerName }</td>
  			</tr>
  		</table>
  	</div>
  	<div><a href="member/loginOut.do" >退出登录</a></div>
  </body>
</html>
