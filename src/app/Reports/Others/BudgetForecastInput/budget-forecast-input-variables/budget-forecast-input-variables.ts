import { Component, HostListener, Input, ElementRef, OnInit, Renderer2, SimpleChanges, ViewChild } from '@angular/core';
import { NgbActiveModal, NgbModal, NgbCalendar, NgbDateParserFormatter, NgbDate, NgbDateStruct, NgbDatepickerConfig, NgbModule } from '@ng-bootstrap/ng-bootstrap';
import { FormsModule, ReactiveFormsModule, FormBuilder, FormArray, FormGroup, Validators } from '@angular/forms';
import { NgxSpinnerModule, NgxSpinnerService } from 'ngx-spinner';
import { ToastrService } from 'ngx-toastr';
import { DatePipe } from '@angular/common';
import { ActivatedRoute, Router } from '@angular/router';
import { BsDatepickerConfig } from 'ngx-bootstrap/datepicker';
import { AnyCatcher } from 'rxjs/internal/AnyCatcher';
import { Title } from '@angular/platform-browser';
import { Api } from '../../../../Core/Providers/Api/api';
import { common } from '../../../../common';
import { SharedModule } from '../../../../Core/Providers/Shared/shared.module';
import { ToastService } from '../../../../Core/Providers/Shared/toast.service';

@Component({
  selector: 'app-budget-forecast-input-variables',
  standalone: true,
  imports: [FormsModule,
    ReactiveFormsModule,
    NgxSpinnerModule, SharedModule, NgbModule],
  templateUrl: './budget-forecast-input-variables.html',
  styleUrl: './budget-forecast-input-variables.scss',
})
export class BudgetForecastInputVariables {
  storeVariableFormGroup: FormGroup;
  stores: any = [];
  storeVarb: any = [];
  actionType = 'N';
  isLoading: boolean = true;

  variableType: any = 'store';

  constructor(private pipe: DatePipe, private service: Api, public comm: common, public activeModal: NgbActiveModal, private ngbmodal: NgbModal, private spinner: NgxSpinnerService, private toast: ToastService, private datepipe: DatePipe, private router: Router, private route: ActivatedRoute,
    private modalservice: NgbModal, public fb: FormBuilder, private title: Title) {

    this.storeVariableFormGroup = fb.group({
      storeval: ['', [Validators.required]],
      expenseVarable: [''],
      storeVarable: ['', [Validators.required]],
      valuetype: ['', [Validators.required]],
      pvrstatus: ['', [Validators.required]],
      category: ['', [Validators.required]],
      formulastatus: [''],
      formula1: [''],
      formula2: [''],
      formula3: [''],
    })

    this.title.setTitle(this.comm.titleName + '-Budget/Forecast  Input');
    const data = {
      title: 'Budget/Forecast  Input',
      path1: '',
      path2: '',
      path3: '',
    };
    this.service.SetHeaderData({ obj: data });
  }


  ngAfterViewInit() {
    this.spinner.show();
    this.getStoresData();
  }

  setVariableType(type: string) {
    this.variableType = type;
    this.getIncomeTargetVariables('71');
  }

  getStoresData() {
    this.service.getStores().subscribe({
      next: (res: any) => {
        console.log('Stores 123:', res);

        if (res?.obj?.storesData) {
          let storeGroup = res.obj.storesData.find((val: any) => val.sg_id == '1');
          if (storeGroup) {
            this.stores = storeGroup.Stores;
            this.getIncomeTargetVariables(this.stores[0]?.ID ?? '71');
          } else {
            console.warn('No store group with sg_id = 1 found');
            this.spinner.hide();
          }
        } else {
          console.warn('No stores data found');
          this.spinner.hide();
        }
      },
      error: (err) => {
        console.error('Error fetching stores:', err);
        this.spinner.hide();
      }
    });
  }

  filteredCategories: any;
  filteredCategoryNames: any;
  filteredFinSummary: any;
  getIncomeTargetVariables(storeId: string) {
    this.spinner.show();
    const obj = {
      StoreID: storeId,
      IV_V_Type: this.variableType
    };

    this.service.postmethod('IncomeTargetVariables/GetIncomeTargetVariablesgetV2', obj)
      .subscribe({
        next: (resp: any) => {
          this.spinner.hide();
          console.log('Resp : ', resp);
          this.filteredCategories = [...new Set(resp.response.map((v: any) => v.Category))];
          this.filteredCategoryNames = [...new Set(resp.response.map((v: any) => v.ITV_CategoryName))];
          this.filteredFinSummary = [...new Set(resp.response.map((v: any) => v.ITV_FinSummary))];
          console.log('Filtered Categories', this.filteredCategories);
          console.log('Filtered CategoryNames', this.filteredCategoryNames);
          console.log('Filtered FinSummary', this.filteredFinSummary);
          let filtered: any[] = resp.response.filter((v: any) => v.ITV_Status !== '');
          let data: any[] = [];
          if (this.variableType === 'store') {
            data = filtered.reduce((r: any, { Category, ...rest }: any) => {
              if (!r.some((o: any) => o.Category === Category)) {
                r.push({
                  Category,
                  ...rest,
                  subdata: filtered
                    .filter((v: any) => v.Category === Category)
                    .sort((a: any, b: any) => a.SequenceNo - b.SequenceNo)
                });
              }
              return r;
            }, []);
          } else if (this.variableType === 'expense') {
            data = filtered.reduce((r: any, { ITV_CategoryName, ...rest }: any) => {
              if (!r.some((o: any) => o.ITV_CategoryName === ITV_CategoryName)) {
                r.push({
                  ITV_CategoryName,
                  ...rest,
                  subdata: filtered
                    .filter((v: any) => v.ITV_CategoryName === ITV_CategoryName)
                    .sort((a: any, b: any) => a.SequenceNo - b.SequenceNo)
                });
              }
              return r;
            }, []);
          }
          data = data.sort((a: any, b: any) => {
            if (a.CategorySequence === b.CategorySequence) {
              return a.SequenceNo - b.SequenceNo;
            }
            return a.CategorySequence - b.CategorySequence;
          });
          this.storeVarb = data;
          console.log('store Varb', this.storeVarb);

        },
        error: (err) => {
          console.error('Error fetching budget variables:', err);
          this.spinner.hide();
        },
        complete: () => {
          this.spinner.hide();
        }
      });
  }





  // GET F(VALIDATIONS)
  get f() {
    return this.storeVariableFormGroup.controls;
  }

  resetForm() {
    this.storeVariableFormGroup.reset({
      storeval: '',
      expenseVarable: '',
      storeVarable: '',
      valuetype: '',
      pvrstatus: '',
      category: '',
      formulastatus: '',
      formula1: '',
      formula2: '',
      formula3: ''
    });
    this.submitted = false;
    // this.addVname = false;
    // this.addSvariable = false;
  }


  activePopover: number = -1;
  @HostListener('document:click', ['$event'])
  onDocumentClick(event: MouseEvent) {
    const clickedInside = (event.target as HTMLElement).closest('.btn-secondary, .canvendar-select-popover');
    if (!clickedInside) {
      this.activePopover = -1;
    }
  }

  getVariableType(typ: any) {
    this.storeVarb = [];
    this.variableType = typ;
    this.activePopover = -1;
    this.getVariablesData();
  }



  change = '0';
  onselect(e: any) {
    this.change = e.target.value;
    console.log('val :', e);
  }

  showFormula: any = false;
  formulaChange(ev: any) {
    console.log(ev.target.checked);
    this.showFormula = ev.target.checked;
  }


  apiResp: any = '';
  showAddForm: boolean = false;
  submitted: boolean = false;
  isEditMode = false;
  editIndex: any = null;
  newVariable: any = {
    Category: '',
    ITV_CategoryName: '',
    ITV_FinSummary: '',
    ITV_Variable_Name: '',
    ITV_Formula: '',
    IV_value_type: 'G',
    ITV_Formula_Status: 'N',
    ITV_Status: 'Y',
    IV_pvr_status: 'N',
    ITV_StoreID: ''
  };

  toggleAddForm() {
    this.showAddForm = !this.showAddForm;
    this.isEditMode = false;
    this.showSequenceChange = false;
    if (!this.showAddForm) {
      this.newVariable = {};
    }
  }

  saveVariable() {
    this.submitted = true;

    if (!this.newVariable.Category || !this.newVariable.ITV_Variable_Name) {
      this.toast.show('Please fill required fields', 'warning', 'Warning');
      return;
    }
    if (this.newVariable.ITV_Formula_Status === 'N') {
      this.newVariable.ITV_Formula = '';
    }
    let Obj = {
      ITV_StoreID: this.newVariable.ITV_StoreID || '71',
      ITV_FinSummary: this.newVariable.ITV_FinSummary || '',
      ITV_Variable_Name: this.newVariable.ITV_Variable_Name,
      Category: this.newVariable.Category,
      ITV_CategoryName: this.newVariable.ITV_CategoryName || '',
      ITV_Formula_Status: this.newVariable.ITV_Formula_Status || 'N',
      ITV_Formula: this.newVariable.ITV_Formula || '',
      ITV_Status: this.newVariable.ITV_Status || 'N',
      ITV_Value_Type: this.newVariable.IV_value_type,
      ITV_Pvr_Status: this.newVariable.IV_pvr_status || 'N'
    };
    console.log('Save Obj :', Obj);
    this.service.postmethod('IncomeTargetVariables', Obj).subscribe({
      next: (res: any) => {
        if (res.status === 200) {
          this.toast.show(res.response[0].message, 'success', 'Success');
          this.getIncomeTargetVariables(this.stores[0]?.ID ?? '71');
          this.toggleAddForm();
        } else {
          this.toast.show('Invalid inputs!', 'danger', 'Error');
        }
        this.submitted = false;
      },
      error: (err) => {
        console.error(err);
        this.toast.show('Something went wrong', 'danger', 'Error');
        this.submitted = false;
      }
    });
  }


  UpdateVariable(item: any, sub: any, ind: any) {
    this.showAddForm = true;       // open form
    this.showSequenceChange = false;
    this.isEditMode = true;        // switch to edit mode
    this.editIndex = { item, sub, ind }; // store reference
    this.newVariable = { ...sub }; // pre-fill form with existing values
  }


  updateVariable() {
    if (this.editIndex) {
      Object.assign(this.editIndex.sub, this.newVariable);

      let Obj = {
        ITV_Id: this.newVariable.ITV_ID,  // ✅
        ITV_StoreID: this.newVariable.ITV_StoreID || 71,
        ITV_FinSummary: this.newVariable.ITV_FinSummary || '',
        ITV_Variable_Name: this.newVariable.ITV_Variable_Name,
        Category: this.newVariable.Category,
        ITV_CategoryName: this.newVariable.ITV_CategoryName || '',
        ITV_Formula_Status: this.newVariable.ITV_Formula_Status || 'N',
        ITV_Formula: this.newVariable.ITV_Formula || '',
        ITV_Status: this.newVariable.ITV_Status || 'N',
        ITV_Value_Type: this.newVariable.IV_value_type,
        ITV_Pvr_Status: this.newVariable.IV_pvr_status || 'N'
      };

      console.log('Update Obj:', Obj);

      this.service.putmethod('IncomeTargetVariables/UpdateIncomeTargetVariablesV1', Obj)
        .subscribe({
          next: (res: any) => {
            if (res.status === 200) {
              this.toast.show(res.response[0].message, 'success', 'Success');
              this.getIncomeTargetVariables(this.stores[0]?.ID ?? '71');
              this.toggleAddForm();
            } else {
              this.toast.show('Invalid inputs!', 'danger', 'Error');
            }
            this.submitted = false;
          },
          error: (err) => {
            console.error(err);
            this.toast.show('Something went wrong', 'danger', 'Error');
            this.submitted = false;
          }
        });
    }
  }





  submitGrid() {
    this.spinner.show();
    if (this.storeVarb && this.storeVarb.length > 0) {
      this.service.putmethod('IncomeTargetVariables/UpdateIncomeTargetBudgetVariableSequence', this.storeVarb).subscribe((res) => {
        console.log('Update Seq : ', res);
        this.spinner.hide();
        this.ngAfterViewInit();
      });
    }
  }


  changeStatus(x: any, ind: any, typ: any) {

    if (typ == 'ITV_Formula') {
      this.storeVarb[ind].formulaStatus = 'Y'
    } else {
      this.storeVarb[ind].changeStatus = 'Y'
    }

  }

  changeStatusFocusOut(x: any, ind: any, typ: any) {
    console.log('focus out : ', x);

    if (typ == 'ITV_Formula') {
      this.storeVarb[ind].formulaStatus = 'N'
    } else {
      this.storeVarb[ind].changeStatus = 'N'
    }

    var obj = {
      "ID": x.ITV_ID,
      "Type": typ,
      "Value": (typ == 'IV_value_type' ? x.IV_value_type : x.ITV_Formula)
    }

    this.service.putmethod('UpdateIncomeTargetBudgetVariables', obj).subscribe((res) => {
      console.log('Update Value Type : ', res);
      this.getVariablesData();
    });
  }


  changePvrStatus(x: any, ind: any, typ: any) {
    this.spinner.show();
    console.log('check it 1: ', x.IV_pvr_status);
    var clickVal = '';

    if (x.IV_pvr_status == 'Y') {
      var clickVal = 'N'
    } else if (x.IV_pvr_status == 'N') {
      var clickVal = 'Y'
    } else if (x.ITV_Status == 'Y') {
      var clickVal = 'N'
    } else {
      var clickVal = 'Y'
    }

    console.log('check it 2: ', clickVal);

    var obj = {
      "ID": x.ITV_Id,
      "Type": typ,
      "Value": clickVal
    }

    this.service.putmethod('UpdateIncomeTargetBudgetVariables', obj).subscribe((res) => {
      console.log('Update Value Type : ', res);
      this.spinner.hide();
      this.getVariablesData();
    });
  }

  cancel() {
    this.actionType = 'N';
    this.getVariablesData();
    this.resetForm();
    //this.router.navigateByUrl('/IncomeTargetVariables-report');  // open welcome component    
  }

  actionTab(typ: any) {
    this.actionType = typ;
    //this.router.navigateByUrl('/IncomeTargetVariablesAdd');  // open welcome component
  }


  getVariablesData() {
    this.spinner.show();
    this.storeVarb = [];
    this.service.getStores().subscribe((res: any) => {

      console.log('Stores :', res)
      if (res.obj.storesData != undefined) {
        this.stores = res.obj.storesData.filter((val: any) => val.sg_id == '1');
        this.stores = this.stores[0].Stores;

      }
    })
    //this.getStoresByGroup();

    var obj = {
      StoreID: "71",
      IV_V_Type: this.variableType
    }


    this.service.postmethod('IncomeTargetVariables/GetIncomeTargetBudgetVariablesget', obj).subscribe((res) => {
      console.log('Resp main : ', res);
      this.storeVarb = res.response;
      this.spinner.hide();
    });
  }

  showSequenceChange: boolean = false;
  showCatSeqChange: boolean = false;
  SeqNoArray: any = [];
  CatSeqNoArray: any = [];
  CategotyName: any;
  CatSeq: any;

  updateCatSeq(Catseq: any) {
    console.log('Cat seq', Catseq);
    this.CatSeqNoArray = Catseq
    this.showSequenceChange = false;
    this.showCatSeqChange = true;
    this.showAddForm = false;
    this.isEditMode = false;
  }
  updateSequenceNo(SeqArray: any, CateName: any, Ind: any) {
    console.log('Seq Array', SeqArray);
    console.log('Cate Name', CateName);
    console.log('Index', Ind);
    this.CatSeq = Ind;
    this.SeqNoArray = SeqArray;
    this.CategotyName = CateName;
    this.showSequenceChange = true;
    this.showCatSeqChange = false;
    this.showAddForm = false;
    this.isEditMode = false;
  }
  draggedIndex: number | null = null;
  onDragStart(event: DragEvent, index: number) {
    this.draggedIndex = index;
    event.dataTransfer?.setData('text/plain', index.toString());
  }
  onDragOver(event: DragEvent) {
    event.preventDefault();
  }
  onDrop(event: DragEvent, targetIndex: number) {
    event.preventDefault();
    if (this.draggedIndex !== null && this.draggedIndex !== targetIndex) {
      const draggedItem = this.SeqNoArray[this.draggedIndex];
      this.SeqNoArray.splice(this.draggedIndex, 1);
      this.SeqNoArray.splice(targetIndex, 0, draggedItem);
      this.updateSequence(); // open modal
    }
    this.draggedIndex = null;
  }
  updateSequence() {
    this.SeqNoArray = this.SeqNoArray.map((item: any, index: number) => ({
      ITV_ID: item.ITV_ID,
      ITV_Variable_Name: item.ITV_Variable_Name,
      SequenceNo: index + 1,
      CategorySequence: this.CatSeq
    }));
  }
  draggedIndexCat: number | null = null;
  onDragStartCat(event: DragEvent, index: number) {
    this.draggedIndexCat = index;
    event.dataTransfer?.setData('text/plain', index.toString());
  }
  onDragOverCat(event: DragEvent) {
    event.preventDefault();
  }
  onDropCat(event: DragEvent, targetIndex: number) {
    event.preventDefault();
    if (this.draggedIndexCat !== null && this.draggedIndexCat !== targetIndex) {
      const draggedItem = this.CatSeqNoArray[this.draggedIndexCat];
      this.CatSeqNoArray.splice(this.draggedIndexCat, 1);
      this.CatSeqNoArray.splice(targetIndex, 0, draggedItem);
      this.updateCatSequence(); // open modal
    }
    this.draggedIndexCat = null;
  }
  updateCatSequence() {
    this.CatSeqNoArray = this.CatSeqNoArray.map((item: any, index: number) => ({
      ITV_FinSummary: item.ITV_FinSummary,
      Category: item.Category,
      ITV_CategoryName: item.ITV_CategoryName,
      CategorySequence: index + 1,
    }));
  }
  saveSequence() {
    if (!this.SeqNoArray || this.SeqNoArray.length === 0) {
      console.warn('Sequence array is empty. Nothing to update.');
      return;
    }
    console.log("New Order:", this.SeqNoArray);
    this.spinner.show();

    this.service.putmethod('IncomeTargetVariables/UpdateVariableSequenceOrder', this.SeqNoArray)
      .subscribe({
        next: (res: any) => {
          console.log('Update Seq:', res);
          this.toast.show('Sequence updated successfully.', 'success', 'Success');
          this.getIncomeTargetVariables(this.stores[0]?.ID ?? '71');
        },
        error: (err: any) => {
          console.error('Error updating sequence:', err);
        },
        complete: () => {
          this.spinner.hide();
        }
      });
  }
  saveCatSequence() {
    if (!this.CatSeqNoArray || this.CatSeqNoArray.length === 0) {
      console.warn('Sequence array is empty. Nothing to update.');
      return;
    }
    console.log("New Order:", this.CatSeqNoArray);
    this.spinner.show();

    this.service.putmethod('IncomeTargetVariables/UpdateVariableCategorySequence', this.CatSeqNoArray)
      .subscribe({
        next: (res: any) => {
          console.log('Update Seq:', res);
          this.toast.show('Category Sequence updated successfully.', 'success', 'Success');
          this.getIncomeTargetVariables(this.stores[0]?.ID ?? '71');
        },
        error: (err: any) => {
          console.error('Error updating sequence:', err);
        },
        complete: () => {
          this.spinner.hide();
        }
      });
  }
  cancelSequence() {
    console.log("Sequence update canceled");
    this.showAddForm = false;
    this.isEditMode = false;
    this.showSequenceChange = false;
    this.showCatSeqChange = false;
  }

}
