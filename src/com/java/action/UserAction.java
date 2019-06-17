package com.java.action;

import java.sql.Connection;
import java.sql.ResultSet;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.struts2.ServletActionContext;

import com.java.dao.UserDao;
import com.java.model.PageBean;
import com.java.model.User;
import com.java.util.DbUtil;
import com.java.util.ExcelUtil;
import com.java.util.JsonUtil;
import com.java.util.ResponseUtil;
import com.java.util.StringUtil;
import com.opensymphony.xwork2.ActionSupport;

public class UserAction extends ActionSupport {

	/**
	 *
	 */
	private static final long serialVersionUID = 1L;

	private String page;
	private String rows;
	private String id;
	private User user;
	private String delId;


	public String getPage() {
		return page;
	}
	public void setPage(String page) {
		this.page = page;
	}
	public String getRows() {
		return rows;
	}
	public void setRows(String rows) {
		this.rows = rows;
	}

	public String getDelId() {
		return delId;
	}
	public void setDelId(String delId) {
		this.delId = delId;
	}
	public User getUser() {
		return user;
	}
	public void setUser(User user) {
		this.user = user;
	}


	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}


	DbUtil dbUtil=new DbUtil();
	UserDao userDao=new UserDao();

	/**
	 * 展示
	 * @return
	 * @throws Exception
	 */
	public String list()throws Exception{
		Connection con=null;
		PageBean pageBean=new PageBean(Integer.parseInt(page),Integer.parseInt(rows));
		try{
			con=dbUtil.getCon();
			JSONObject result=new JSONObject();
			JSONArray jsonArray=JsonUtil.formatRsToJsonArray(userDao.userList(con, pageBean));
			int total=userDao.userCount(con);
			result.put("rows", jsonArray);
			result.put("total", total);
			ResponseUtil.write(ServletActionContext.getResponse(),result);
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			try {
				dbUtil.closeCon(con);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return null;
	}

	/**
	 * 添加
	 * @return
	 * @throws Exception
	 */
	public String save()throws Exception{
		if(StringUtil.isNotEmpty(id)){
			user.setId(Integer.parseInt(id));
		}
		Connection con=null;
		try{
			con=dbUtil.getCon();
			int saveNums=0;
			JSONObject result=new JSONObject();
			if(StringUtil.isNotEmpty(id)){
				saveNums=userDao.userModify(con, user);
			}else{
				saveNums=userDao.userAdd(con, user);
			}
			if(saveNums>0){
				result.put("success", "true");
			}else{
				result.put("success", "true");
				result.put("errorMsg", "保存失败");
			}
			ResponseUtil.write(ServletActionContext.getResponse(), result);
		}catch(Exception e){
			e.printStackTrace();
		}finally{
			try {
				dbUtil.closeCon(con);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return null;
	}

	/**
	 * 删除数据
	 * @return
	 * @throws Exception
	 */
	public String delete()throws Exception{
		Connection con=null;
		try {
			con=dbUtil.getCon();
			JSONObject result=new JSONObject();
			int delNums=userDao.userDelete(con, delId);
			if(delNums==1){
				result.put("success", "true");
			}else{
				result.put("errorMsg", "删除失败");
			}
			ResponseUtil.write(ServletActionContext.getResponse(), result);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				dbUtil.closeCon(con);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return null;
	}


	public String export()throws Exception{
		Connection con=null;
		try {
			con=dbUtil.getCon();
			Workbook wb=new HSSFWorkbook();  //创建工作簿

			String headers[]={"编号","姓名","电话","Email","QQ"};
			ResultSet rs=userDao.userList(con, null);
			ExcelUtil.fillExcelData(rs, wb, headers);
			ResponseUtil.export(ServletActionContext.getResponse(), wb, "导出excel.xls");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				dbUtil.closeCon(con);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		return null;
	}

	/**
	 * 用模板导出后台
	 * @return
	 * @throws Exception
	 */
	public String export2() throws  Exception{
		Connection con=null;
		try {
			con = dbUtil.getCon();
			ResultSet rs=userDao.userList(con, null);
			Workbook wb = ExcelUtil.fillExcelDataWithTemplate(userDao.userList(con, null), "userExporTemplate.xls");
			ResponseUtil.export(ServletActionContext.getResponse(), wb, "模板导出excel.xls");

		}catch (Exception e){
			e.printStackTrace();
		}finally {
			try {
			dbUtil.closeCon(con);
		}catch (Exception e){
			e.printStackTrace();
		}

		}
		return null;
	}



}
