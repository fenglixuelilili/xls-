<template>
<div @drop="_drop" @dragenter="_suppress" @dragover="_suppress">

  <div class="row">
    <div class="col-xs-12">
      <form class="form-inline">
        <div class="form-group">
          <label for="file" class="labels">
            点击上传
            <input type="file" class="form-control" id="file" :accept="SheetJSFT" @change="_change" />
          </label>
        </div>
      </form>
    </div>
  </div>

  <div class="row">
    <div class="col-xs-12">
      <el-button type="primary" :disabled="data.length ? false : true" class="btn btn-success" @click="_export">导出excel</el-button>
    </div>
  </div>

  <div class="el-table el-table--fit el-table--enable-row-hover el-table--enable-row-transition" style="width: 100%;">
    <div class="hidden-columns">
      <div></div>
      <div></div>
      <div></div>
    </div>

    <div class="el-table__header-wrapper">
      <table cellspacing="0" cellpadding="0" border="0" class="el-table__header" style="width: 1278px;">
        <colgroup>
          <col name="el-table_2_column_4" width="180">
          <col name="el-table_2_column_5" width="180">
          <col name="el-table_2_column_6" width="918">    
          <col name="gutter" width="0">
        </colgroup>
        <thead class="has-gutter">
          <tr v-for="(r, key) in data" :key="key" v-if="key === 0">
            <th colspan="1" rowspan="1" class="el-table_2_column_4     is-leaf" v-for="c in cols" :key="c.key" @click="handle(c.key)" contenteditable="true">
              <div class="cell">{{r[c.key]}}</div>
            </th>
          </tr>
        </thead>
      </table>
    </div>

    <div class="el-table__body-wrapper is-scrolling-none">
      <table cellspacing="0" cellpadding="0" border="0" class="el-table__body" style="width: 100%;">
        <colgroup>
        <col name="el-table_2_column_4" width="180">
        <col name="el-table_2_column_4" width="180">
        <col name="el-table_2_column_4" width="180">
        </colgroup>
        <tbody>
          <!-- <tr v-for="(r, key) in data" :key="key">
            <td v-for="c in cols" :key="c.key" @click="handle(c.key)" contenteditable="true"> {{ r[c.key] }}</td> 
          </tr> -->
          <tr class="el-table__row" v-for="(r, key) in data" :key="key" v-if="key !== 0">
            <td rowspan="1" colspan="1" class="el-table_2_column_4" v-for="c in cols" :key="c.key" @click="handle(c.key)" contenteditable="true">
              <div class="cell"> {{ r[c.key] }}</div>
            </td>
          </tr><!---->
        </tbody>
      </table><!----><!---->
    </div><!----><!----><!----><!---->
    <div class="el-table__column-resize-proxy" style="display: none;"></div>
  </div>

  <div class="row">
    <div class="col-xs-12">
    </div>
  </div>
















  <div id="app" v-cloak>
        <input type="file" @change="importExcel($event.target)" />
        <div style="overflow: auto;" v-if="tableTbody&&tableTbody.length>0">
            <table class="table table-bordered" style="min-width: 100%;">
                <thead>
                    <tr>
                        <th>#</th>
                        <th v-for="(item,index) in tableHeader" :key="index">
                            {{item}}
                        </th>
                    </tr>
                </thead>
                <tbody>
                    <tr v-for="(row,index) in tableTbody" :key="index">
                        <th scope="row">{{index}}</th>
                        <td v-for="col in tableHeader">{{row[col]}}</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </div>
</div>

</template>

<script>
import XLSX from 'xlsx'
const makecols = refstr => Array(XLSX.utils.decode_range(refstr).e.c + 1).fill(0).map((x, i) => ({name: XLSX.utils.encode_col(i), key:i}));

const _SheetJSFT = ['xlsx', 'xlsb', 'xlsm', 'xls', 'xml', 'csv', 'txt', 'ods', 'fods', 'uos', 'sylk', 'dif', 'dbf', 'prn', 'qpw', '123', 'wb*', 'wq*', 'html', 'htm'
].map(function (x) { return '.' + x }).join(',')
export default {
  data () {
    return {
      data: ['SheetJS'.split(''), '1234567'.split('')],
      cols: [
        {name: 't', key: 0},
        {name: 'e', key: 1},
        {name: 's', key: 2},
        {name: 't', key: 3}
      ],
      SheetJSFT: _SheetJSFT,
      tableData: [{
            date: '2016-05-02',
            name: '王小虎',
            address: '上海市普陀区金沙江路 1518 弄'
          }, {
            date: '2016-05-04',
            name: '王小虎',
            address: '上海市普陀区金沙江路 1517 弄'
          }, {
            date: '2016-05-01',
            name: '王小虎',
            address: '上海市普陀区金沙江路 1519 弄'
          }, {
            date: '2016-05-03',
            name: '王小虎',
            address: '上海市普陀区金沙江路 1516 弄'
          }]
    }
  },
  methods: {
    _suppress (evt) { evt.stopPropagation(); evt.preventDefault() },
    _drop (evt) {
      evt.stopPropagation(); evt.preventDefault()
      const files = evt.dataTransfer.files
      if (files && files[0]) this._file(files[0])
    },
    _change (evt) {
      const files = evt.target.files
      if (files && files[0]) this._file(files[0])
    },
    _export (evt) {
      /* convert state to workbook */
      const ws = XLSX.utils.aoa_to_sheet(this.data)
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, ws, 'SheetJS')
      /* generate file and send to client */
      XLSX.writeFile(wb, 'sheetjs.xlsx')
    },
    _file (file) {
      /* Boilerplate to set up FileReader */
      const reader = new FileReader()
      reader.onload = (e) => {
        /* Parse data */
        const bstr = e.target.result
        const wb = XLSX.read(bstr, {type: 'binary'})
        /* Get first worksheet */
        const wsname = wb.SheetNames[0]
        const ws = wb.Sheets[wsname]
        /* Convert array of arrays */
        const data = XLSX.utils.sheet_to_json(ws, {header:1})
        /* Update state */
        this.data = data
        this.cols = makecols(ws['!ref'])
      }
      reader.readAsBinaryString(file)
    },
    handle (index) {
      console.log(this.cols[index].key)
    },




    importExcel(obj) {
                if (!obj.files) {
                    return;
                }
                let file = obj.files[0],
                    types = file.name.split('.')[1],
                    fileType = ["xlsx", "xlc", "xlm", "xls", "xlt", "xlw", "csv"].some(item => item === types);
                if (!fileType) {
                    alert("格式错误！请重新选择");
                    return;
                }
                this.file2Xce(file).then(tabJson => {
                    if (tabJson && tabJson.length > 0) {
                        this.tableHeader = Object.keys(tabJson[0]);
                        this.tableTbody = tabJson;
                    }
                });
            },
            file2Xce(file) {
                return new Promise(function (resolve, reject) {
                    let reader = new FileReader();
                    reader.onload = function (e) {
                        let data = e.target.result;
                        this.wb = XLSX.read(data, {
                            type: 'binary'
                        });
                        resolve(XLSX.utils.sheet_to_json(this.wb.Sheets[this.wb.SheetNames[0]]));
                    };
                    reader.readAsBinaryString(file);
                });
            }
        }
  
}
</script>
<style>
  #file {
    display: none
  }
  .labels {
    /* width: 100px; */
    font-size: 14px;
    padding: 6px 10px;
    border-radius: 5px;
    background: #409EFF;
    color: #FFF;
    border: 1px solid #409EFF;
    position: absolute;
    top: 2%;
    left: 40%;
  }
</style>
