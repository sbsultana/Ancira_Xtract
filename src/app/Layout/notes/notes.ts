import { Component, OnInit, Input, Renderer2, Output, EventEmitter, } from '@angular/core';
import { NgbActiveModal, NgbModal } from '@ng-bootstrap/ng-bootstrap';
import { Router } from '@angular/router'
import { Sharedservice } from '../../Core/Providers/Shared/sharedservice';
import { SharedModule } from '../../Core/Providers/Shared/shared.module';
import { ToastService } from '../../Core/Providers/Shared/toast.service';



@Component({
  selector: 'app-notes',
  imports: [SharedModule],
  standalone: true,
  templateUrl: './notes.html',
  styleUrl: './notes.scss'
})
export class Notes {
  @Input() notesData: any;
  @Output() onClose = new EventEmitter();
  @Output() SavedNotesData = new EventEmitter();

  // @Output() SavedNotesData = new EventEmitter();

  history: any[] = [];

  userid: any = ''
  spinnerLoader: boolean = false;
  notes: any = '';
  menu: any = []
  curl: any = '';
  renderer: any;
  FullData: any;
  menuData: any = []
  constructor(public shared: Sharedservice, private router: Router, private toast: ToastService) {
    // this.renderer.listen('window', 'click', (e: Event) => {
    //   const TagName = e.target as HTMLButtonElement;
    // });
    this.curl = router.url;
  }

  ngOnInit(): void {
    this.userid = JSON.parse(localStorage.getItem('userInfo')!)?.user_Info;
    this.menuData = this.shared.common.menuData.flat();
    console.log(this.menuData, 'Menu Data');
    this.menu = this.menuData.filter((v: any) => v.mod_filename.indexOf(this.curl) > 0)[0].mod_id

    this.history = this.notesData?.history || [];
    this.FullData = this.notesData?.fulldata
    console.log('FullData', this.FullData, '........................................................');

    if (this.FullData) {

      // ✅ Normalize COMMENT
      if (typeof this.FullData.COMMENT === 'string') {
        try {
          this.FullData.COMMENT = JSON.parse(this.FullData.COMMENT);
        } catch {
          this.FullData.COMMENT = [];
        }
      }


    }

  }


  get hasHistory(): boolean {
    return this.notesData?.history?.length > 0;
  }



  close(val: any) {
    // this.ngbmodel.dismissAll();
    this.onClose.emit('C');

  }

  savenotes() {
    if (this.notes == '') {
      this.toast.show('Please enter notes', 'warning', 'Warning');
    } else {
      let obj = {}
      if (this.notesData.apiRoute == 'AddGeneralNotes') {
        obj = {
          "GN_AS_ID": this.notesData.store,
          "GN_Title1": this.notesData.title1,
          "GN_Title2": this.notesData.title2,
          "GN_Text": this.notes,
          "GN_NS_ID": '',
          "GN_Module_ID":this.menu,
          "GN_Active": "Y",
          "GN_Module_Code":this.notesData.moduleCode,
          "GN_UserId": JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid,
        };
      } else if (this.notesData.apiRoute == 'AddNotesAction') {
        obj = {
          "AS_ID": this.notesData.store,
          "Title": this.notesData.mainkey,
          "Module":  this.notesData.module,
          "Notes": this.notes,
          'UserID': JSON.parse(localStorage.getItem('userInfo')!)?.user_Info?.userid,
        };
      }

      this.spinnerLoader = true
      this.shared.api.postmethod(this.shared.common.routeEndpoint + this.notesData.apiRoute, obj).subscribe(
        (res) => {
          if (res.status == 200) {
            this.toast.show('Notes inserted successfully', 'success', 'Success')
            let data={notes:this.notes+'- ' +JSON.parse(localStorage.getItem('userInfo')!)?.fullName +' - ('+ this.shared.datePipe.transform(new Date(),'MM/dd/yy') + ')' }

            this.SavedNotesData.emit(data);
            this.onClose.emit('S');
            // this.ngbmodel.dismissAll();

          } else {
            this.toast.show('Invalid Details', 'danger', 'Error');
          }
        },
        (error) => { }
      );
    }
  }







}
