<template>
  <div class="app">
    <Upload action="" :before-upload="beforeUpload">
      <Button icon="ios-cloud-upload-outline">导入文件</Button>
    </Upload>

    <Table
      border
      height="600"
      :loading="loading"
      :columns="columns"
      :data="departments"
    />
  </div>
</template>

<script setup>
import { ref, nextTick } from "vue";
import { read, utils } from "xlsx";
import { Button, Table, Upload } from "view-ui-plus";
import { useDepartmentDetail } from "./composition/departmentDetail";

const { sheetData, columns, departments } = useDepartmentDetail();

const loading = ref(false);

function beforeUpload(file) {
  const fr = new FileReader();

  sheetData.length = 0;

  loading.value = true;

  if (!file) {
    return alert("文件不对");
  }

  fr.onload = (e) => {
    const data = e.target.result,
      { SheetNames, Sheets } = read(data, { type: "buffer" }),
      sheetName = SheetNames[SheetNames.length - 1],
      sheet = Sheets[sheetName],
      json = utils.sheet_to_json(sheet, { header: 1, defval: null });

    sheetData.push(...json);

    nextTick(() => {
      loading.value = false;
    });
  };

  fr.readAsArrayBuffer(file);

  return false;
}
</script>

<style scoped>
.app {
  padding: 16px;
  box-sizing: border-box;
}
::v-deep(.choice-list) {
  list-style: none;
}
</style>
