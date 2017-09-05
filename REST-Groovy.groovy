@Grapes([
  @Grab(group='io.rest-assured', module='rest-assured', version='3.0.1'),
  @Grab(group='org.apache.poi', module='poi', version='3.16'),
  @Grab(group='org.apache.poi', module='poi-ooxml', version='3.16')
])


import io.restassured.RestAssured
import static io.restassured.RestAssured.*
import io.restassured.http.Header
import io.restassured.response.Response
import io.restassured.builder.RequestSpecBuilder
import io.restassured.specification.RequestSpecification
import groovy.json.*
import org.apache.poi.ss.usermodel.*
import java.io.File
import groovy.sql.*
import groovy.json.*

File f = new File("C:\\Users\\clinton.paiva\\Desktop\\testexcel.xls");

def excelreader=WorkbookFactory.create(f,null,true)
println "Has ${excelreader.getNumberOfSheets()} sheets"
Sheet secondSheet = excelreader.getSheetAt(1);
Iterator<Row> iterator1 = secondSheet.iterator();
def list1=[]
while (iterator1.hasNext()) {
	Row nextRow = iterator1.next();
	Iterator<Cell> cellIterator = nextRow.cellIterator();
	while (cellIterator.hasNext()) {
		Cell cell = cellIterator.next();
		list1.add(cell.getStringCellValue());
	}
}
Sheet thirdSheet = excelreader.getSheetAt(2);
Iterator<Row> iterator2 = thirdSheet.iterator();
def list2=[]
while (iterator2.hasNext()) {
	Row nextRow = iterator2.next();
	Iterator<Cell> cellIterator = nextRow.cellIterator();
	while (cellIterator.hasNext()) {
		Cell cell = cellIterator.next();
		list2.add(cell.getStringCellValue());
	}
}
Sheet firstSheet = excelreader.getSheetAt(0);
Iterator<Row> iterator = firstSheet.iterator();
def list=[],excelcount=0
while (iterator.hasNext()) {
	excelcount++
	println "Executing Test Case " +excelcount
	list.clear();
	Row nextRow = iterator.next();
    Iterator<Cell> cellIterator = nextRow.cellIterator();
    while (cellIterator.hasNext()) {
		Cell cell = cellIterator.next();
		list.add(cell.getStringCellValue());
	}
	
	//Reading Instance Details
	def username = list1[1]
	def password = list1[2]
	def url = list1[0]
	
	def endpoint = list[0]
	def responseCode=list[1].replace("\"", "").toInteger();
	def jsontest=list[3]
	RestAssured.useRelaxedHTTPSValidation()
	Header header = new Header("Content-Type", "application/x-www-form-urlencoded")
	
	// Login to the instance
	
	Response response = given().formParam("username", username).formParam("password",password).header(header).request().post(url + "auth/basic")
	RequestSpecBuilder builder = new RequestSpecBuilder()
	//builder.addHeader("X-Csrf-Token", response.getCookie("Csrf-Token"))
	builder.addCookies([ "JSESSION_ID" : response.getCookie("JSESSION_ID") ])
	builder.setContentType("application/json; charset=UTF-8")
	builder.addHeader("accept", "application/json")
	requestSpec = builder.build()
	
	// Make a REST API Call
	println "Executing for EndPoint: "+endpoint
	Response r = given().spec(requestSpec).when().get(url + "api/7.0/" + endpoint)
	println "Return Status Code = " + r.getStatusCode()
	assert r.getStatusCode()==responseCode
	println "Return JSON = " + r.asString()
	def jsonRes=new JsonSlurper().parseText(r.asString())
	
	//Splitting JSON string in excel
	def count1=0
	
	def key = jsontest.split("\\.")
	def var= key[0]
	def no = key.size()
	
	def els = jsonRes."$var"
	//Checking count of occurrence of JSON value
	for(int i=1;i<no;i++) {
		var = key[i]
		els = els."$var".each{count1++}
	}
	println count1
	
	//Reading DB details
	url= list2[0]
	username =  list2[1]
	password =  list2[2]
	driver =  list2[3]
	query=list[2]
	if(query==null){
		break;
	}else{
	// Groovy Sql connection test
		sql = Sql.newInstance(url, username, password, driver)
		def data
		try {
			//Executing JSON query from excel
			data =sql.rows(query)
			assert count1 == data.count[0]
		 } catch (AssertionError e) {
			 println 'Assertion Failed Expected '+count1 +' but found in db ' + data.count[0]
		 }
	}
}

