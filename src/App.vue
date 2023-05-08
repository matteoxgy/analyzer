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
import { computed, nextTick, reactive, ref } from "vue";
import { utils, read } from "xlsx";
import { Button, Table, Upload } from "view-ui-plus";

const sheetData = reactive([]),
  questions = computed(
    () => sheetData[0]?.filter((c, index) => 10 <= index && index <= 19) || []
  ),
  columns = computed(() => {
    let cols = [
      {
        title: "部门",
        key: "department",
        fixed: "left",
        width: 200,
      },
    ];

    cols.push(
      ...questions.value.map((question, index) => ({
        title: question,
        key: index,
        width: 200,
        render: (h, { row }) =>
          h(
            "ul",
            {
              class: "choice-list",
            },
            row[index].map((choice) =>
              h("li", null, `${choice[0]}:${choice[1]}`)
            )
          ),
      }))
    );

    return cols;
  }),
  departments = computed(() => {
    const res = {};

    sheetData.slice(1).forEach((row) => {
      const depName = row[6],
        answers = row
          .slice(10)
          .map((answer) =>
            (answer.match(/[A-Z]\./g) || []).map((c) => c.charAt(0))
          )
          .filter((answer) => answer !== "(");

      let depStatus = res[depName];

      // 没有记录过该部门
      if (!depStatus) {
        depStatus = [];

        res[depName] = depStatus;
      }

      answers.forEach((ans, index) => {
        let answerStatus = depStatus[index];

        // 没有记录过该问题
        if (!answerStatus) {
          answerStatus = {};

          depStatus[index] = answerStatus;
        }

        ans.forEach((answer) => {
          // 没有记录过该回答
          if (!answerStatus[answer]) {
            answerStatus[answer] = 0;
          }

          answerStatus[answer]++;
        });
      });
    });

    return Object.entries(res).map((item) => {
      const [department, depData] = item;

      const out = {
        department,
      };

      depData.forEach((s, index) => {
        out[index] = sortAttributes(s);
      });

      return out;
    });
  }),
  loading = ref(false);

function sortAttributes(answers) {
  return Object.entries(answers).sort(
    (a, b) => a[0].charCodeAt(0) - b[0].charCodeAt(0)
  );
}

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
