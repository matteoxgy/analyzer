import { read, utils } from "xlsx";

const uploaderEl = document.getElementById("file-uploader"),
  listEl = document.getElementById("list"),
  tableEl = document.getElementById("table"),
  tableHeader = document.getElementById("table-header");

uploaderEl.addEventListener("change", (event) => {
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

    stats(json);
  };

  fr.readAsArrayBuffer(file);
});

function stats(sheet) {
  sheet = sheet.map((row) =>
    row.filter((c, index) => index === 6 || (10 <= index && index <= 19))
  );

  const questions = sheet[0].slice(1),
    res = new Map(questions.map((question) => [question, new Map()]));

  questions.forEach((question) => {
    const th = document.createElement("th");

    th.innerText = question;

    tableHeader.appendChild(th);
  });

  // 遍历每一行
  for (let i = 1; i < sheet.length; i++) {
    const row = sheet[i],
      dep = row[0];

    console.log(dep);

    // 遍历每一列
    row.forEach((col, index) => {
      if (index === 0) {
        return;
      }

      const question = questions[index - 1],
        detail = res.get(question),
        answer = col.charAt(0);

      if (detail.has(answer)) {
        detail.set(answer, detail.get(answer) + 1);
      } else {
        detail.set(answer, 1);
      }
    });
  }

  const listFragment = document.createDocumentFragment();

  Array.from(res).forEach((ele) => {
    const [question, detail] = ele,
      li = document.createElement("li"),
      title = document.createElement("h3"),
      detailList = document.createElement("ul");

    Array.from(detail).forEach((item) => {
      const [answer, count] = item,
        el = document.createElement("li");

      el.innerText = `选择${answer}的人有${count}个`;

      detailList.appendChild(el);
    });

    title.innerText = question;

    li.appendChild(title);
    li.appendChild(detailList);

    listFragment.appendChild(li);
  });

  listEl.appendChild(listFragment);
}
