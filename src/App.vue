<template>
  <button type="button" @click="exportXLSX">Export</button>
</template>

<script lang="ts">
import { defineComponent } from "vue";
import XLSX from "xlsx";
import XLSXS from "xlsx-style";
import FileSaver from "file-saver";

const data = [
  ["查詢區間 2021/1/1 起 2021/1/31 迄"],
  ["組別", "派案數", "已結案", "未結案", "逾期已結案", "逾期未結案"],
  ["第一組", 2, 3, 4, 5, 0],
];

//单元格外侧框线
const borderAll = {
  top: {
    style: "thin",
  },
  bottom: {
    style: "thin",
  },
  left: {
    style: "thin",
  },
  right: {
    style: "thin",
  },
};
// 寬度
const wscols = [
  { wch: 10 },
  { wch: 10 },
  { wch: 10 },
  { wch: 10 },
  { wch: 10 },
  { wch: 10 },
];

// 首列合併
var merge = { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } };

export default defineComponent({
  setup() {
    const exportXLSX = () => {
      var filename = "write.xlsx";
      var ws_name = "SheetJS";

      const wb = XLSX.utils.book_new();
      const _ws = XLSX.utils.aoa_to_sheet(data);
      const ws = setWS(_ws);

      /* add worksheet to workbook */
      XLSX.utils.book_append_sheet(wb, ws, ws_name);

      /* write workbook */
      var result = XLSXS.write(wb, {
        type: "buffer",
      });

      FileSaver.saveAs(
        new Blob([result], { type: "application/octet-stream" }),
        filename
      );
    };

    const setWS = (ws: XLSX.WorkSheet) => {
      return setColRow(setStyle(ws));
    };

    const setColRow = (ws: XLSX.WorkSheet) => {
      // 設定寬度
      ws["!cols"] = wscols;
      // 合併首列
      if (!ws["!merges"]) ws["!merges"] = [];
      ws["!merges"].push(merge);
      return ws;
    };

    const setStyle = (ws: XLSX.WorkSheet) => {
      for (let key in ws) {
        if (ws[key] instanceof Object) {
          ws[key].s = {
            border: borderAll,
            wch: 30,
          };
          // 表頭格式化
          if (key === "A1") {
            ws[key].s.font = { sz: 15 };
            ws[key].s.font = { sz: 15 };
            ws[key].s.alignment = {
              horizontal: "center", //水平居中对齐
              vertical: "center",
            };
          }
        }
      }
      return ws;
    };

    return {
      exportXLSX,
    };
  },
});
</script>
