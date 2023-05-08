<template>
  <div>
    <input type="file" @change="onFileChange" />

    <div class="table-wrapper">
      <table class="table" border="2px">
        <tr class="table-head">
          <th></th>
          <th v-for="question of questions" :key="question">{{ question }}</th>
        </tr>

        <tr
          class="dep-row"
          v-for="(answers, department) in departments"
          :class="department === active ? 'active' : ''"
          :title="department"
          @click="onRowClick(department)"
        >
          <td>{{ department }}</td>
          <td v-for="answer of sort(answers)" :key="answer">
            <ul v-for="choice of answer">
              {{
                choice
              }}
            </ul>
          </td>
        </tr>
      </table>
    </div>
  </div>
</template>

<script setup>
import { utils, read } from "xlsx";
import { ref, computed, reactive } from "vue";

const sheetData = reactive([]),
  questions = computed(
    () => sheetData[0]?.filter((c, index) => 10 <= index && index <= 19) || []
  ),
  departments = computed(() => {
    const res = {};

    sheetData.slice(1).forEach((row) => {
      const depName = row[6],
        answers = row
          .slice(10)
          .map((answer) => answer.match(/[A-Z]/g) || [])
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

    return res;
  }),
  active = ref("");

function onFileChange(event) {
  const file = event.target.files[0],
    fr = new FileReader();

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
  };

  fr.readAsArrayBuffer(file);
}

function onRowClick(department) {
  active.value = department;
}

function sort(answers) {
  return answers.map((answerDetail) =>
    Object.entries(answerDetail).sort(
      (a, b) => a[0].charCodeAt(0) - b[0].charCodeAt(0)
    )
  );
}
</script>

<style scoped>
.table-wrapper {
  overflow: auto;
}
.table {
  width: 4000px;
}
.table-head {
  background-color: #f2f2f2;
}
.dep-row.active {
  background-color: #f2f2f2;
}
</style>
