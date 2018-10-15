import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import 'jquery';
import 'bluebird';


import * as strings from 'ImportExcelWebpartWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';


export interface IImportExcelWebpartWebPartProps {
  description: string;
}

export default class ImportExcelWebpartWebPart extends BaseClientSideWebPart<IImportExcelWebpartWebPartProps> {
  private _currentWebUrl: string;

  constructor() {
    super();

  }

  public render(): void {
    this._currentWebUrl = this.context.pageContext.web.absoluteUrl;
    localStorage.setItem("url", this._currentWebUrl);
    require("./app/UploadCenter.css");
    require("./app/ImportExcel.js");
  
    this.domElement.innerHTML = `
    <div id="Loader" class="overlay">
      <div class="box">
        <div class="loader8"></div>
      </div>
      </div>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css"/>
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css"/>
    <div id="myModal" class="modal fade" role="dialog">
  <div class="modal-dialog">
        <!-- Modal content-->
        <div class="modal-content">
          <div class="modal-body">
            <p>Data successfully imported!</p>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-default" data-dismiss="modal">Close</button>
          </div>
        </div>
      </div>
    </div>
    <div class="row">
    <p class="header"><img class="img-responsive header_logo" src="https://46dtbf3k4dl51vghpj6qqocj-wpengine.netdna-ssl.com/wp-content/themes/ilac/img/ILAC-white-logo.png">​​</p>​
  </div>
  <div class="col-md-12">
    <div class="panel panel-default">    
      <div class="panel-body" id="form">
        <!--  Details -->
        <div class="form-group" style="margin-bottom:0px;">
          <div class="row" id="ddlRow">
            <div class="controls col-sm-12">
              <select id="ddtype1">
                <option value="0">Select List</option>
                <option value="Students">Students</option>
                  <option value="Teachers">Teachers</option>
                  <option value="Classes">Classes</option>
              </select>
              <label class="active" for="ddtype1">List<span style="color:red;">*</span><span style="display:none;" class="Error" id="spnListError">Select List.</span></label>
            </div>
          </div>
          <div class="row" style="display:none;" id="rowLocation">
            <div class="controls col-sm-12">
              <select id="ddlLocation">
              </select>
              <label class="active" for="ddtype1">Location<span style="color:red;">*</span><span style="display:none;" class="Error" id="spnLocationError">Select Location.</span></label>
            </div>
          </div>
          <!-- Row 1 Ends -->
                  <!-- Row 2 -->
            <div class="row" style="margin-top:15px; margin-bottom:15px;">
              <div class="controls col-sm-12">
			  <label for="txtTfile_inputitle" id="file_input_lable">File Upload<span style="color:red;">*</span><span style="display:none;" class=	"Error" id="spnFileError">Select Excel file(.xlsx || .xls).</span></label>
                <input type="file" id="excelfile" class="floatLabel"/>
              </div>
            </div>
          <!-- Row 2 Ends -->
          <!-- Buttons -->
          <div class="row">
            <div class="controls col-sm-3">
              <button type="button" id="viewfile"><i class="fa fa-file-excel-o" aria-hidden="true"></i>&nbsp; Import</button>
                      </div>
            <div class="controls col-sm-3">
              <button type="button" id="btnClear"><i class="fa fa-times" aria-hidden="true"></i></i>&nbsp; Cancel</button>
            </div>
          </div>
        </div>
    </div>
  </div>`;
  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
