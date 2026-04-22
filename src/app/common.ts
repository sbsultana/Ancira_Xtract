import { Injectable } from '@angular/core';
// Update the import to the correct path for ApiService
import { Api } from '../app/Core/Providers/Api/api';
import { environment } from '../environments/environment';
@Injectable({
  providedIn: 'root'
})
export class common {
  public groupsandstores: any = [];
  public groupsandstoresAll: any = [];

  public userId: any = 0;
  public pageName: any = '';
  public titleName: any = 'Ancira';
  public excelName: any = 'Ancira';
  public solutionurl: any = 'https://fredbeans.axelautomotive.com/'
  public axelonedashboard: any = 'https://fredbeans.axelautomotive.com/dashboard'
  public logo: any = 'images/Logo.png'
  public firststore: any = [{ "storename": "FORD OF DOYLESTOWN", "ID": 1, "storecode": "FB-A", "sg_id": 1, "sg_name": "Fred Beans " }]
  public menuUrl: any = 'https://fbxtract.axelautomotive.com/';
  public redirectionFrom: any = '';
  public menuData: any = [];
  public routeEndpoint: any = 'Ancira/'
  public reconID: any = 7;
  public allstores: any = [];
  public OtherStoresData: any = []
  public otherstoreids: any = [];
  public DefaultOtherstoresSelection: any = ''
  public modules: any = '';
  public currentRoute: any = ''
  public localUserData: any = {
    "userid": 1,
    "UserName": "Prasad Chavali",
    "fname": "Prasad",
    "lname": "Chavali",
    "mailid": "prasad.chavali@axelautomotive.com",
    "phno": "9999999999",
    "roleid": 100,
    "profilepic": "",
    "userstores": "1000,8,7,4,35,1,32,40,50,18,31,3,2,17,41,42,51,12,25,15,9,60,63,62,61,64,65,67,66,5,14,30,11",
    "Store_Ids": "71,53,8,7,4,35,1,32,40,50,25,18,31,3,70,72,2,17,41,55,42,51,12,73,54,9,15,5,14,30,11",
    "status": "Y",
    "title": "Super Admin",
    "uid": 1,
    "Xtract": 1,
    "Touch": 1,
    "Xperience": 1,
    "Xchange": 3,
    "Xiom": 1,
    "Tracs": 1,
    "EmpID": ''
  }
  public ReconStores: any = []
  public AccountingReconStores: any = []
  public ReconServiceLaborStores: any = []
  public MobileServiceGL: any = [];
  public completeUserDetails: any = {}



  widgets: any;
  products: any;
  adtoken: any = "";
  btnback: boolean = false;
  username: any;
  userObj: any;
  storename: any;
  empname: any;
  empmail: any;
  preinfo: any = '';
  caseno: any;
  lastupdate: any = "";
  //   modules:any=[];
  xtrctlink: any;
  mtd: any = 0;
  userInfo: any;
  constructor(public apiService: Api) {
    let user: any = localStorage.getItem("userobj");
    this.userObj = (user != '') ? JSON.parse(user) : [];

    const userinfo: any = localStorage.getItem('userInfo');
    console.log(localStorage.getItem('userInfo'));

    const userInfo: any = JSON.parse(userinfo);
    this.userInfo = userInfo;


  }
  navtoapp(id: any) {
    let user: any = localStorage.getItem("userobj");
    console.log('Commaon User : ', user);
    this.userObj = (user != '') ? JSON.parse(user) : [];
    var url = environment.apiUrl;

    this.apiService.getsessionid(this.userInfo.user_aou_AD_userid, this.userInfo.GU_URL).subscribe(
      (data: any) => {
        let token =
        {
          "userid": this.userInfo.user_aou_AD_userid,
          "role": this.userInfo.user_Info.roleid,
          "session": data.response,
          "store": this.userInfo.user_Info.ustores,
          "flag": "M",
          "groupid": (localStorage.getItem('grpId') != '' && localStorage.getItem('grpId') != null) ? localStorage.getItem('grpId') : "1",
        };
        let tkn = btoa(JSON.stringify(token));
        if (id == 2) {
          window.open("https://cndrxtract.axelautomotive.com/?token=" + tkn, "_blank");
        }
        if (id == 5) {
          window.open("https://cndrxtract.axelautomotive.com/SalesGross?token=" + tkn, "_blank");
        }
        if (id == 6) {
          window.open("https://cndrxtract.axelautomotive.com/ServiceGross?token=" + tkn, "_blank");
        }
        if (id == 7) {
          window.open("https://cndrxtract.axelautomotive.com/PartsGross?token=" + tkn, "_blank");

        }
        if (id == 8) {
          window.open("https://cndrxtract.axelautomotive.com/Floorplan?token=" + tkn, "_blank");

        }
        if (id == 9) {
          window.open("https://cndrxtract.axelautomotive.com/SalesPersonRanking?token=" + tkn, "_blank");

        }
        if (id == 11) {
          window.open("https://cndrxtract.axelautomotive.com/ServiceAdvisorRanking?token=" + tkn, "_blank");

        }

        if (id == 10) {
          window.open("https://cndrxtract.axelautomotive.com/FandIManagerRanking?token=" + tkn, "_blank");

          //FandIMangerRanking
        }

      });
  }
}