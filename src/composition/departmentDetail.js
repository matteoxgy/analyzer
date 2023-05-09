import { computed, reactive } from "vue";

export function useDepartmentDetail() {
  const sheetData = reactive([]),
    questions = computed(
      () => sheetData[0]?.filter((c, index) => 10 <= index && index <= 19) || []
    ),
    departments = computed(() => {
      const depSituation = {
        总计: [],
      };

      sheetData.slice(1).forEach((row) => {
        const depName = row[6],
          answers = row
            .slice(10)
            .map((answer) =>
              (answer.match(/[A-Z]\./g) || []).map((c) => c.charAt(0))
            )
            .filter((answer) => answer !== "(");

        let depStatus = depSituation[depName];

        // 没有记录过该部门
        if (!depStatus) {
          depStatus = [];

          depSituation[depName] = depStatus;
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

            // 加入总计行
            const totalQuestionSituation = depSituation["总计"];

            let curQuestionTotal = totalQuestionSituation[index];

            // 总计行没有记录过此题
            if (!curQuestionTotal) {
              curQuestionTotal = {};

              totalQuestionSituation[index] = curQuestionTotal;
            }

            // 没有加过此回答
            if (!curQuestionTotal[answer]) {
              curQuestionTotal[answer] = 0;
            }

            curQuestionTotal[answer]++;
          });
        });
      });

      const res = Object.entries(depSituation).map((item) => {
        const [department, depData] = item;

        const out = {
          department,
        };

        depData.forEach((s, index) => {
          out[index] = sortAttributes(s);
        });

        return out;
      });

      return res;
    }),
    columns = computed(() => {
      let cols = [
        {
          title: "部门",
          key: "department",
          fixed: "left",
          width: 200,
          filters: departments.value.map((department) => ({
            label: department.department,
            value: department.department,
          })),
          filterMultiple: true,
          filterMethod(value, row) {
            return row.department === value;
          },
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
              (row[index] || []).map((choice) =>
                h("li", null, `${choice[0]}:${choice[1]}`)
              )
            ),
        }))
      );

      return cols;
    });

  function sortAttributes(answers) {
    return Object.entries(answers).sort(
      (a, b) => a[0].charCodeAt(0) - b[0].charCodeAt(0)
    );
  }

  return {
    sheetData,
    columns,
    questions,
    departments,
  };
}
