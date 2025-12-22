var FileName = ".xlsx";
// xlsx objects
var Workbook, Spent, Future, Over, Plan, SST;
// html objects
var Spending, Upcoming, Charts, Plans, Formulas;
// default start spent row
var StartSpent = 4;
// encode due to xlsx issues
var Pound = decodeURIComponent("%C2%A3");

// get overrideable min 1 year history based on 1st april start
var Beginning = new Date();
Beginning.setFullYear(Beginning.getFullYear() - (Beginning.getMonth() >= 3 ? 1 : 2));
Beginning.setDate(1);
Beginning.setMonth(3);

// get overrideable today point
var Today = new Date();

// get year end name
var Year = Today.getFullYear() + (Today.getMonth() + 1 >= Beginning.getMonth() + 1 ? 1 : 0);

// functions

async function log(text) {
  document.getElementById('output').innerHTML += (text + '<br>');
  await timeout();
}

function HideUnhide(ths) {
  if (ths.parentNode.children[0].checked) {
    ths.parentNode.children[0].checked = false;
    event.preventDefault();
  }
  setTimeout(function () {
    document.getElementsByTagName('p')[0].scrollIntoView({ block: 'end', behavior: 'smooth' });
  }, 500);
}

function Highlight (ths) {
  if (ths.style.backgroundColor)
    ths.style.backgroundColor = null;
  else
    ths.style.backgroundColor = "#ddd";
}

function NewRow(ths, data = []) {
  var m = ths.parentNode.getAttribute('max')?.toString();
  if (!m)
    m = '0';
  var k = parseInt(m) + 1;
  ths.parentNode.setAttribute('max', k);
  
  if (!data)
    data = [];
  if (data === false || data === true)
    data = [ data ];

  var row = ths.cloneNode(true);
  rename(row,0,0,k,data[0]);
  rename(row,1,0,k,data[1]);
  rename(row,2,0,k,data[2]);
  rename(row,3,0,k,data[3]);
  rename(row,4,0,k,data[4]);
  rename(row,5,0,k,data[5]);
  rename(row,5,1,k);
  rename(row,5,2,k);
  rename(row,6,0,k,data[6]);
  rename(row,7,0,k,data[7]);
  rename(row,8,0,k);
  rename(row,9,0,k);
  ths.before(row);
}

function rename(row, col, obj, k, value = null) {
  if (!row.children[col])
    return;
  if (!row.children[col].children[obj])
    return;
  row.children[col].children[obj].name = row.children[col].children[obj].name.replace(/[0-9]*/g, '') + k;
  if (row.children[col].children[obj].type == "checkbox")
    row.children[col].children[obj].checked = value ? true : null;
  else if (row.children[col].children[obj].type != "button")
    row.children[col].children[obj].value = value === undefined ? null : value;
}

function DelRow(ths) {
  try {
    if (ths.parentNode.children.length == 1)
      NewRow(ths, true);
    ths.parentNode.removeChild(ths);
  } catch { }
}

function clearRows(ths) {
  while (ths.length > 1) {
    DelRow(ths[0]);
  }
  DelRow(ths[0]);
}

async function Edit(t) {
  // await so highlight triggers
  await timeout();
  
  if (!confirm("Remove "+t.value.trim()+" to reedit?"))
    return;

  t.disabled = true;
  await log('Editing');

  var row = parseInt(t.value.trim());
  var dt = new Date(t.parentNode.parentNode.children[1].innerText.trim());
  var val = parseFloat(t.parentNode.parentNode.children[2].innerText.replace(Pound, "").replace(',', '').trim());
  var desc = t.parentNode.parentNode.children[3].innerText.trim();
  t.parentNode.parentNode.remove();

  if (row > StartSpent)
    StartSpent = row + 1;

  await log('...reducing');

  var spent = Spent.getRange("A4:Z" + Spent.getMaxRows()).getValues();
  spent.splice(row - 4, 1);
  Spent.getRange("A4:Z" + Spent.getMaxRows()).clear();
  Spent.getRange("A4:Z" + (3 + spent.length)).setValues(spent);

  await log('...formulas');

  var oldCc = (Spent.getRange("D3:D3").getFormula() ?? "").replace(/^=/, "");
  var newCc = oldCc.split("-").filter(c => c != "D" + row).join("-");
  var hadCc = oldCc != newCc;

  var cc = newCc.split("D");
  var a = "=" + cc[0];
  for (let i = 1; i < cc.length; i++) {
    if (row <= parseInt(cc[i]))
      a += "D" + (parseInt(cc[i]) - 1) + cc[i].replace(/^[0-9]+/, '');
    else
      a += "D" + cc[i];
  }

  Spent.getRange("D3:D3").setFormula(a);

  Spent.getRange("D2:D2").setFormula("=SUM(D3:D" + (3 + spent.length) + ")-D1+E3");

  var top = Spent.getRange("F2:Z2").getFormulasR1C1();
  for (let i = 0; i < top[0].length; i++) {
    try {
      let m = top[0][i].match(/\:([A-Za-z])([0-9]+)/);
      let f = parseInt(m[2]);
      if (f >= row)
        top[0][i] = top[0][i].replace(m[0], ":" + m[1] + (f - 1));
    } catch (e) { }
  }
  Spent.getRange("F2:Z2").setFormulasR1C1(top);

  await log('...editable');

  if (Spending[0].children[3].children[0].value)
    NewRow(Spending[0]);
  Spending[0].children[0].children[0].checked = hadCc;
  Spending[0].children[1].children[0].value = dt.toISOString().substring(0, 10);
  Spending[0].children[2].children[0].value = val.toFixed(2);
  Spending[0].children[3].children[0].value = desc;

  await log('...calc');

  calc();
  await setupSpentFields();
  await Filter(false);

  ths.disabled = null;
  await log('...done');
}

async function IncludeSpent(ths) {
  ths.disabled = true;
  await log('Including');

  await setupSpentFields();

  await log('...inserting');

  var spent = Spent.getRange("A4:Z" + Spent.getMaxRows()).getValues();

  for (let n = Spending.length - 1; n >= 0; n--) {
    let row = Spending[n];
    if (!row.children[3].children[0].value)
      continue;

    let v = (typeof row.children[1].children[0].value == "string" ? row.children[1].children[0].value : row.children[1].children[0].value.toISOString().substring(0, 10)).split('-')

    for (var p = 0; p < spent.length; p++)
      if (new Date(spent[p][0], spent[p][1] - 1, spent[p][2]) <= new Date(v[0], v[1] - 1, v[2]))
        break;

    var cc = (Spent.getRange("D3:D3").getFormula() ?? "").replace(/^=/, "").split("D");
    var a = "=" + cc[0];
    for (let i = 1; i < cc.length; i++) {
      if ((p + 4) <= parseInt(cc[i]))
        a += "D" + (parseInt(cc[i]) + 1) + cc[i].replace(/^[0-9]+/, '');
      else
        a += "D" + cc[i];
    }

    if (row.children[0].children[0].checked)
      a += "-D" + (4 + p);
    Spent.getRange("D3:D3").setFormula(a);

    spent.splice(p, 0, [
      parseInt(v[0]),
      parseInt(v[1]),
      parseInt(v[2]),
      parseFloat(row.children[2].children[0].value),
      row.children[3].children[0].value.trim()
    ]);

    var top = Spent.getRange("F2:Z2").getFormulasR1C1();
    for (let i = 0; i < top[0].length; i++) {
      try {
        let m = top[0][i].match(/\:([A-Za-z])([0-9]+)/);
        let f = parseInt(m[2]);
        if (f > (p + 4))
          top[0][i] = top[0][i].replace(m[0], ":" + m[1] + (f + 1));
      } catch (e) { }
    }
    Spent.getRange("F2:Z2").setFormulasR1C1(top);

    // move start along for populate spent to work
    if (4 + n + p > StartSpent)
      StartSpent = 4 + n + p + 1;

    DelRow(row);
  }

  await log('...setting formulas');

  // clear off for populate spent to work
  for (let m = StartSpent; m >= 4; m--)
    Spent.getRange("F" + m + ":F" + m).setFormula("");

  Spent.getRange("A4:Z" + (3 + spent.length)).setValues(spent);

  Spent.getRange("D2:D2").setFormula("=SUM(D3:D" + (3 + spent.length) + ")-D1+E3");

  await log('...calc');

  calc();
  await setupSpentFields();
  await Filter(true);

  ths.disabled = null;
  await log('...done');
}

var debounce;

function Search(ths) {
  if (!ths || !ths.value || ths.value.trim().length < 3) {
    var uls = document.getElementsByTagName("ul");
    for (let u = 0; u < uls.length; u++)
      uls[u].style.display = 'none';
    return;
  }

  clearTimeout(debounce);
  debounce = setTimeout(() => {
    var p = SST.filter(s => s.toLowerCase().includes(ths.value.toLowerCase())).slice(0, 10);
    ths.nextElementSibling.innerHTML = p.map(s => `<li onclick="Select(this, '${s}')">${s}</li>`).join('');
    ths.nextElementSibling.innerHTML += `<li style="text-align: center" onclick="Select(this, '')">Close</li>`;
    ths.nextElementSibling.style.display = p.length == 0 ? 'none' : 'block';
  }, 1000);
}

function FilterSearch() {
  clearTimeout(debounce);
  debounce = setTimeout(() => {
    Filter();
  }, 1000);
}

function Select(ths, val) {
  if (val)
    ths.parentNode.previousElementSibling.value = val;
  ths.parentNode.style.display = 'none';
}

async function Filter(incFuture) {
  await log('Filtering');

  var template = `
    <tr onclick='Highlight(this)'>
      <td style=width:22px>{5}</td>
      <td style=width:96px>{0}</td>
      <td style="color: {4}">{1}</td>
      <td {6}>{2}</td>
    </tr>
      `;

  var s = document.getElementById('spent');
  s.innerHTML = "";
  
  await log('...setting spent');

  var range = (document.getElementById("range") ?? {})["value"] ?? "99";
  var cols = parseInt((document.getElementById("cols") ?? {})["value"] ?? "-1");
  var filter = ((document.getElementById("filter") ?? {})["value"] ?? "").toLowerCase();
  var max = Spent.getMaxRows() - 1;
  if (range > max)
    range = max;
  var filtering = cols >= 0 || !!filter;
  var subtotal = 0.0;

  var spent = Spent.getRange("A4:Z" + range).getValues();
  for (let i in spent)
    if (spent[i][3] !== null && spent[i][3] !== ''
      && spent[i][4] !== null && spent[i][4] !== ''
      && (cols < 0 || spent[i][cols] || spent[i][cols] === '-')
      && (!filter || ~spent[i][4].toLowerCase().indexOf(filter))
    ) {
      s.innerHTML += template
        .replace('{5}', "<input type='button' onclick='Edit(this)' value='" + (parseInt(i) + 4).toString() + "'/>")
        .replace('{0}', !spent[i][2] ? '' : !spent[i][0] ? spent[i][2] : `${spent[i][0]}-${(spent[i][1] < 10 ? "0" : "") + spent[i][1]}-${(spent[i][2] < 10 ? "0" : "") + spent[i][2]}`)
        .replace('{1}', `${(spent[i][3] || 0).toLocaleString("en-GB", { style: "currency", currency: "GBP" })}`)
        .replace('{2}', `${spent[i][4]}`)
        .replace('{4}', spent[i][3] < 0 ? 'red' : 'inherited');
      if (filtering)
        subtotal += spent[i][3] || 0;
    }

  if (filtering)
    s.innerHTML += template
      .replace('{5}', "")
      .replace('{0}', "")
      .replace('{1}', `${subtotal.toLocaleString("en-GB", { style: "currency", currency: "GBP" })}`)
      .replace('{2}', "Subtotal")
      .replace('{4}', subtotal < 0 ? 'red' : 'inherited');
  
  SST = [];
  Spent.getRange("E4:E999").getValues().forEach(s => {
    let sst = (s[0] || "").trim();
    if (sst && !SST.includes(sst))
      SST.push(sst);
  });
  SST = SST.sort((a, b) => a.length - b.length);

  await log('...setting formulas');

  document.getElementById('D1').value = Spent.getRange("D1:D1").getValue();
  document.getElementById('D2').value = Spent.getRange("D2:D2").getValue().toFixed(2);
  document.getElementById('Cc').value = Spent.getRange("D3:D3").getValue().toFixed(2);
  document.getElementById('D3').value = Spent.getRange("D3:D3").getFormula();
  document.getElementById('E3').value = Spent.getRange("E3:E3").getValue();
  document.getElementById('Sum').value = (parseFloat(Spent.getRange("E3:E3").getValue()) + parseFloat(Spent.getRange("D3:D3").getValue())).toFixed(2);

  if (!incFuture)
    return;

  await log('...setting over');

  var o = document.getElementById('over').getElementsByTagName('table')[0];
  o.innerHTML = template
    .replace('{5}', 'Y-M')
    .replace('{0}', 'Under')
    .replace('{1}', 'Annual')
    .replace('{2}', 'Low')
    .replace('{4}', 'inherited')

  var over = Over.getRange("A2:J99").getValues();
  for (let i in over)
    if (over[i][0])
      o.innerHTML += template
        .replace('{5}', over[i][0] + "-" + over[i][1])
        .replace('{0}', !over[i][5] ? '' : over[i][5].toLocaleString("en-GB", { style: "currency", currency: "GBP" }))
        .replace('{1}', !over[i][8] ? '' : over[i][8].toLocaleString("en-GB", { style: "currency", currency: "GBP" }))
        .replace('{4}', over[i][8] < 0 ? 'red' : 'inherited')
        .replace('{2}', !over[i][9] ? '' : over[i][9].toLocaleString("en-GB", { style: "currency", currency: "GBP" }));

  o.innerHTML += '<tr><td colspan=4><hr></td></tr>' + template
    .replace('{5}', 'Y-M-D')
    .replace('{0}', 'Value')
    .replace('{1}', 'Balance')
    .replace('{4}', 'inherited')
    .replace('{2}', '')

  var c = 0;
  var F1 = Workbook.Sheets['F1'] || Workbook.Sheets['F2'] || Workbook.Sheets['F3'] || Workbook.Sheets['F4'] || Workbook.Sheets['F5'] || Future;
  var f1 = F1.getRange("A2:G" + F1.getMaxRows()).getValues();
  for (let i in f1)
    if (f1[i][6]) {
      o.innerHTML += template
        .replace('{5}', f1[i][0] + "-" + (f1[i][1] < 10 ? '0' : '') + f1[i][1] + "-" + (f1[i][2] < 10 ? '0' : '') + f1[i][2])
        .replace('{0}', !f1[i][3] ? '' : f1[i][3].toLocaleString("en-GB", { style: "currency", currency: "GBP" }))
        .replace('{1}', !f1[i][6] ? '' : f1[i][6].toLocaleString("en-GB", { style: "currency", currency: "GBP" }))
        .replace('{4}', f1[i][6] < 0 ? 'red' : 'inherited')
        .replace('{2}', f1[i][4]);
      c++;
      if (c > 100)
        break;
    }
  
  await log('...setting future');

  var futs = document.getElementById('futures').getElementsByTagName('table')[0];
  futs.innerHTML = template
          .replace('{5}', 'Balance')
          .replace('{0}', 'Date')
          .replace('{1}', 'Value')
          .replace('{2}', 'Desc');
  
  clearRows(Spending);
  clearRows(Upcoming);

  var future = Future.getRange("A2:G" + Future.getMaxRows()).getValues()
  var next = new Date(Today);
  //next.setDate(next.getDate()+1+(next.getDay()==5?1:0));
  next.setHours(13);
  var week = new Date(next);
  week.setDate(week.getDate() + 7);
  var len = parseInt(Spent.getRange("G1:G1").getValue() ?? "3");
  var last = new Date(Year + len, 0, 0);
  var end = new Date(Year + 1, 0, 0);

  var c = 10;
  var foundBlank = false;
  for (let f = 0; f < future.length; f++) {
    if (!foundBlank)
      foundBlank = (future[f][4] || "") == "";
    else {
      let d = new Date(future[f][0], future[f][1] - 1, future[f][2]);
      d.setHours(12);

      if (d <= next) {
        Spending[0].children[0].children[0].checked = null;
        Spending[0].children[1].children[0].value = d.toISOString().substring(0, 10);
        Spending[0].children[2].children[0].value = future[f][3] ? future[f][3].toFixed(2) : 0;
        Spending[0].children[3].children[0].value = future[f][4];

        NewRow(Spending[0]);
      } else if (d <= last) {
        let tr = document.createElement('tr');
        tr.innerHTML = template.replace('<tr>', '').replace('</tr>', '')
          .replace('{5}', '')
          .replace('{0}', d.toISOString().substring(0, 10))
          .replace('{1}', (future[f][3] || 0).toLocaleString("en-GB", { style: "currency", currency: "GBP" }))
          .replace('{2}', future[f][4])
          .replace('{6}', 'colspan=3')
          .replace('{4}', future[f][3] < 0 ? 'red' : 'inherited')
        Upcoming[0].before(tr);

        c--;
        if (c <= 0 && d >= week) {
          last = d;
        }
      } else if (d <= end) {
        futs.innerHTML += template
            .replace('{5}', future[f][6])
            .replace('{0}', d.toISOString().substring(0, 10))
            .replace('{1}', (future[f][3] || 0).toLocaleString("en-GB", { style: "currency", currency: "GBP" }))
            .replace('{2}', future[f][4])
            .replace('{4}', future[f][3] < 0 ? 'red' : 'inherited');
      } else
        break;
    }
  }

  // cleanup additional added row
  DelRow(Spending[0]);

  future = !Workbook.Sheets['F1'] ? [] : Workbook.Sheets['F1'].getRange("A2:E" + Workbook.Sheets['F1'].getMaxRows()).getValues();
  next.setMonth(next.getMonth() + 1);

  var foundBlank = false;
  for (let f = 0; f < future.length; f++) {
    if ((future[f][4] || "") == "")
      foundBlank = true;
    else if (foundBlank) {
      let d = new Date(future[f][0], future[f][1] - 1, future[f][2]);
      d.setHours(12);

      if (d > next) {
        break;
      }

      if (d > last) {
        let tr = document.createElement('tr');
        tr.innerHTML = template.replace('<tr>', '').replace('</tr>', '')
          .replace('{5}', '')
          .replace('{0}', d.toISOString().substring(0, 10))
          .replace('{1}', (future[f][3] || 0).toLocaleString("en-GB", { style: "currency", currency: "GBP" }))
          .replace('{2}', future[f][4])
          .replace('{6}', 'colspan=3')
          .replace('{4}', future[f][3] < 0 ? 'red' : 'inherited')
        Upcoming[0].before(tr);
      }
    }
  }

  var c = Future.getRange("F2:G" + Future.getMaxRows()).getValues();
  var l = c.filter(r => r[0] && (r[1] || r[1] === 0)).map(r => {
    return (typeof r[0] == 'number' ? new Date((r[0] - (25567 + 1))*86400*1000) : r[0]).toJSON().substring(0,10)
  });
  var d = c.filter(r => r[0] && (r[1] || r[1] === 0)).map(r => r[1]);
  Charts[0].style.display = 'block';
  new Chart(Charts[0], {
   type: 'line',
   data: {
    labels: l,
    datasets: [{
     label: 'Balance: ' + Charts.length,
     data: d,
    }]
   },
   options: {
    plugins: {
     zoom: {
      pan: {
       enabled: true,
       mode: 'x'
      },
      zoom: {
       wheel: {
        enabled: true,
       },
       pinch: {
        enabled: true
       }
      }
     }
    }
   }
  });
  Charts[0].before(document.createElement("canvas"));
  Charts[0].style.display = 'none';
}

async function PopulateFormulas(ths) {
  clearRows(Plans);
  clearRows(Formulas);
  
  await log('Getting formulas');

  document.getElementById("formulas").style.display = 'block';
  ths.nextElementSibling.style.display = 'initial';
  ths.nextElementSibling.nextElementSibling.style.display = "initial";
  ths.style.display = 'none';

  var headings = Spent.getRange("F1:Z1").getValues();
  var formulas = Spent.getRange("F2:Z3").getFormulasR1C1();

  var col = 'Z';
  for (let c = headings[0].length - 1; c >= 0; c--) {
    NewRow(Formulas[0], [col, headings[0][c], formulas[0][c], formulas[1][c]]);
    col = XLSX.utils.encode_col(XLSX.utils.decode_col(col) - 1)
  }
  DelRow(Formulas[Formulas.length - 1]);
}

async function SaveFormulas(ths) {
  if (Formulas.length == 1 && !row.children[0].children[0].value)
    return;

  await log('Formulating');

  document.getElementById("formulas").style.display = 'none';
  ths.previousElementSibling.style.display = 'initial';
  ths.nextElementSibling.style.display = "none";
  ths.style.display = 'none';

  var headings = [];
  var totals = [];
  var formulas = [];
  for (let n = 0; n < Formulas.length; n++) {
    let row = Formulas[n];

    let h = row.children[1].children[0].value;
    let t = row.children[2].children[0].value;
    let f = row.children[3].children[0].value;

    headings.push(h);
    totals.push(t);
    formulas.push(f);

    let m = t.match(/:[A-Za-z]+([0-9]+)\)/g);
    // replace as match group not working
    if (m && parseInt(m[0].replace(/[^0-9]/g,'')) > StartSpent)
      StartSpent = parseInt(m[0].replace(/[^0-9]/g,''));
  }

  Spent.getRange("F1:Z1").setValues([headings]);
  Spent.getRange("F2:Z3").setFormulasR1C1([totals, formulas]);

  clearRows(Plans);
  clearRows(Formulas);
  
  await log('...calc');
  calc();

  setupSpentFields();
  Filter(true);

  await log('...done');
}

async function SortPlan(sort) {
  await log('Sorting plans');
  calc();
  
  var plan = [];
  var today = Today.toJSON().substring(0, 10);
  for (let n = 0; n < Plans.length; n++) {
    let row = Plans[n];

    let f = row.children[5].children[0].type == "number"
        ? parseFloat(row.children[5].children[0].value || '0')
        : parseFloat(row.children[5].children[2].value || '0');
    let v = row.children[5].children[0].type == "text" 
      ? row.children[5].children[0].value
      : row.children[5].children[2].value;
    
    let p = [
      !row.children[0].children[0].value ? null : parseInt(row.children[0].children[0].value),
      !row.children[1].children[0].value ? null : parseInt(row.children[1].children[0].value),
      !row.children[2].children[0].value ? null : parseInt(row.children[2].children[0].value),
      !row.children[3].children[0].value ? null : typeof row.children[3].children[0].value == "string" ? row.children[3].children[0].value : row.children[3].children[0].value.toISOString(),
      !row.children[4].children[0].value ? null : typeof row.children[4].children[0].value == "string" ? row.children[4].children[0].value : row.children[4].children[0].value.toISOString(),
      null,
      row.children[6].children[0].value.trim(),
      row.children[7].children[0].value === null || row.children[7].children[0].value === '' ? null : parseInt(row.children[7].children[0].value),
      f,
      (v || "").trim().replace(/^=+/g, ""),
      null
    ];

    if ((p[0] || p[1] || p[2]) && (!p[3] || p[3] > today) && (p[4] < today))
      p[10] = p[8]/(p[0]?p[0]:1)*12/(p[1]?p[1]:12)*365/(p[2]?p[2]:365);

    plan.push(p);
  }

  let direction = sort < 0 ? -1 : 1;
  sort = sort * direction;
  
  plan.sort((a,b) => {
    let aS = a[sort];
    let bS = b[sort];
    if (sort == 0) {
      aS = 10_000*aS + 100*a[1] + a[2];
      bS = 10_000*bS + 100*b[1] + b[2];
    }
    if (aS === null && bS === null)
      return 0;
    if (aS === null)
      return 1;
    if (bS === null)
      return -1;
    if (aS < bS)
      return -1 * direction;
    if (aS > bS)
      return 1 * direction;
    if (aS = bS)
      return 0;
  });

  await log('...outputting');

  clearRows(Plans);

  for (let p = plan.length - 1; p >= 0; p--) {
    NewRow(Plans[0], plan[p]);
    // repush due to datatype
    Plans[0].children[5].children[0].type = "number";
    Plans[0].children[5].children[2].type = "text";
    Plans[0].children[5].children[0].value = plan[p][8];
    Plans[0].children[5].children[2].value = '='+plan[p][9];
    if (plan[p][9] != "")
      ToggleType(Plans[0].children[5].children[0], Plans[0].children[5].children[2]);
  }
  DelRow(Plans[Plans.length - 1]);

  await log('...done');
}

async function PopulatePlan(ths) {
  clearRows(Plans);
  clearRows(Formulas);
  
  await log('Getting plans');

  document.getElementById("plans").style.display = 'block';
  ths.nextElementSibling.style.display = 'initial';
  ths.nextElementSibling.nextElementSibling.style.display = "initial";
  ths.style.display = 'none';

  var plan = Plan.getRange("A2:L" + Plan.getMaxRows()).getValues();
  var planf = Plan.getRange("J2:J" + (plan.length + 1)).getFormulasR1C1();
  
  for (let p = plan.length - 1; p >= 0; p--) {
    Plans[0].children[0].children[0].value = plan[p][0];
    Plans[0].children[1].children[0].value = plan[p][1];
    Plans[0].children[2].children[0].value = plan[p][2];
    if (plan[p][3] > 1900) {
      let s = new Date(plan[p][3], plan[p][4] - 1, plan[p][5]);
      s.setHours(12);
      Plans[0].children[3].children[0].value = s.toISOString().substring(0, 10);
    }
    if (plan[p][6] > 1900) {
      let e = new Date(plan[p][6], plan[p][7] - 1, plan[p][8]);
      e.setHours(12);
      Plans[0].children[4].children[0].value = e.toISOString().substring(0, 10);
    }
    Plans[0].children[5].children[0].type = "number";
    Plans[0].children[5].children[2].type = "text";
    Plans[0].children[5].children[0].value = plan[p][9];
    if ((planf[p][0] || "").trim().replace(/^=+/g, "") != "") {
      Plans[0].children[5].children[2].value = "=" + planf[p][0];
      ToggleType(Plans[0].children[5].children[0], Plans[0].children[5].children[2])
    }
    Plans[0].children[6].children[0].value = plan[p][10];
    Plans[0].children[7].children[0].value = plan[p][11];

    if (p > 0)
      NewRow(Plans[0]);
  }

  await log('...done');
}

function ToggleType (ths, nxt) {
  let n = nxt.value;
  let p = ths.value;
  nxt.type = ths.type;
  ths.type = ths.type == 'number' ? 'text' : 'number';
  ths.value = n;
  nxt.value = p;
}

function Cancel(ths) {
  clearRows(Plans);
  clearRows(Formulas);
  
  document.getElementById("plans").style.display = 'none';
  document.getElementById("formulas").style.display = 'none';
  ths.previousElementSibling.style.display = 'none';
  ths.previousElementSibling.previousElementSibling.style.display = "initial";
  ths.style.display = 'none';
}

async function SavePlan(ths) {
  if (Plans.length == 1 && !row.children[5].children[0].value)
    return;

  await log('Planning');

  document.getElementById("plans").style.display = 'none';
  ths.previousElementSibling.style.display = 'initial';
  ths.nextElementSibling.style.display = "none";
  ths.style.display = 'none';

  var plan = [];
  for (let n = 0; n < Plans.length; n++) {
    let row = Plans[n];

    let s = !row.children[3].children[0].value ? [null, null, null] : (typeof row.children[3].children[0].value == "string" ? row.children[3].children[0].value : row.children[3].children[0].value.toISOString().substring(0, 10)).split('-')
    let e = !row.children[4].children[0].value ? [null, null, null] : (typeof row.children[4].children[0].value == "string" ? row.children[4].children[0].value : row.children[4].children[0].value.toISOString().substring(0, 10)).split('-')
    let f = row.children[5].children[0].type == "number"
        ? parseFloat(row.children[5].children[0].value || '0')
        : parseFloat(row.children[5].children[2].value || '0');
    
    plan.push([
      !row.children[0].children[0].value ? null : parseInt(row.children[0].children[0].value),
      !row.children[1].children[0].value ? null : parseInt(row.children[1].children[0].value),
      !row.children[2].children[0].value ? null : parseInt(row.children[2].children[0].value),
      s[2] ? parseInt(s[0]) : null,
      parseInt(s[1]),
      parseInt(s[2]),
      e[2] ? parseInt(e[0]) : null,
      parseInt(e[1]),
      parseInt(e[2]),
      f,
      row.children[6].children[0].value.trim(),
      row.children[7].children[0].value === null || row.children[7].children[0].value === '' ? null : parseInt(row.children[7].children[0].value)
    ]);
  }

  Plan.getRange("A2:L" + Plan.getMaxRows()).clear();
  Plan.getRange("A2:L" + (plan.length + 1)).setValues(plan);

  await log('...formulas');

  var formulas = Plan.getRange("M2:O2").getFormulasR1C1()[0];
  for (let n = 0; n < Plans.length; n++) {
    let row = Plans[n];

    let v = row.children[5].children[0].type == "text" 
      ? row.children[5].children[0].value
      : row.children[5].children[2].value;
    
    if ((v || "").trim().replace(/^=+/g, "") != '')
      Plan.getRange("J" + (n + 2) + ":J" + (n + 2)).setFormulasR1C1([[v]]);

    let f = [];
    f[0] = formulas[0].replace(/([A-Za-z]+)2/g, '$1' + (n + 2));
    f[1] = formulas[1].replace(/([A-Za-z]+)2/g, '$1' + (n + 2));
    f[2] = formulas[2].replace(/([A-Za-z]+)2/g, '$1' + (n + 2));

    if (row.children[6].children[0].value != null && row.children[6].children[0].value != '')
      Plan.getRange("M" + (n + 2) + ":O" + (n + 2)).setFormulasR1C1([f]);
  }

  clearRows(Plans);
  clearRows(Formulas);
  
  await log('...calc');

  calc();

  await log('...done');
}

async function UpdateCC() {
  await log('Overwriting');

  Spent.getRange("D1:D1").setValue(parseFloat(document.getElementById('D1').value));
  Spent.getRange("D3:D3").setFormula(document.getElementById('D3').value);
  Spent.getRange("E3:E3").setValue(parseFloat(document.getElementById('E3').value));

  calc();
  await setupSpentFields();
  await Filter(true);

  await log('...done');
}

function StartChange(ths) {
  if (!ths || ths == null || ths == "")
    Beginning = new Date(2011, 3, 1);
  else {
    var val = typeof ths == "string" ? ths : ths.toISOString();
    Beginning = new Date(val.substring(0, 4), parseInt(val.substring(5, 7)) - 1, val.substring(8, 10), 12, 0, 0);
  }
  Year = Today.getFullYear() + (Today.getMonth() + 1 >= Beginning.getMonth() + 1 ? 1 : 0);
}

function TodayChange(ths) {
  if (!ths || ths == null || ths == "")
    Today = new Date();
  else {
    var val = typeof ths == "string" ? ths : ths.toISOString();
    Today = new Date(val.substring(0, 4), parseInt(val.substring(5, 7)) - 1, val.substring(8, 10), 12, 0, 0);
  }
  Year = Today.getFullYear() + (Today.getMonth() + 1 >= Beginning.getMonth() + 1 ? 1 : 0);
}

function Upload() {
  document.getElementById('File').value = null;
  document.getElementById('File').click();
}

async function FileChange(ths) {
  if (ths.files.length != 1)
    return document.getElementById('Upload').disabled = null;

  ths.previousElementSibling.style.display = 'initial';

  document.getElementById('output').innerHTML = ('Loading...<br>');

  document.getElementById('future').style.display = "none";
  document.getElementById('over').style.display = "none";
  document.getElementById('options').style.display = "none";

  var f = ths.files[0];
  FileName = f.name;

  await log(`...loading ${FileName}`);

  var reader = new FileReader();
  reader.onerror = function () {
    alert("File reading error " + FileName);
  };
  reader.onload = fileLoaded
  reader.readAsArrayBuffer(f);
}

async function RunAll() {
  await setupOneYear();
  await timeout();

  await setupLimited();
  await timeout();

  await setupAllYears();
  await timeout();

  await setupSpentFields();
  await Filter(true);
  await timeout();

  await Saving();
}

async function setupSpentFields() {
  await log('...blanks');

  // find blank year or calcs in spent to copy top row into
  var row = StartSpent;
  while (Spent.getRange("A" + row + ":A" + row).getValue() == "" && row < Spent.getMaxRows()) {
    row++;
  }

  while ((Spent.getRange("F" + row + ":F" + row).getFormula() ?? "") == "" && Spent.getRange("F" + row + ":Z" + row).getValues()[0].map(function (c) { return c || '' }).join('') == "" && row < Spent.getMaxRows()) {
    row++;
  }

  log('...start '+row);

  // get latest date below blanks
  var last = Spent.getRange("A" + row + ":C" + row).getValues();
  if (last[0][0] == "")
    last[0] = [Beginning.getFullYear(), Beginning.getMonth() + 1, Beginning.getDate()];
  var sum = Spent.getRange("F3:Z3").getFormulasR1C1();
  var y = last[0][0];
  var m = last[0][1];
  var d = last[0][2];

  // set dates correctly and formulas for before today only to allow while above to continue operating
  var daterow = row - 1;
  while (daterow >= 4) {
    if (Spent.getRange("A" + daterow + ":A" + daterow).getValue() == "") {
      var c = Spent.getRange("C" + daterow + ":C" + daterow).getValue();

      // new month
      if (c < d) {
        m++;
        if (m > 12) {
          m = 1;
          y++;
        }
      }

      d = c;

      Spent.getRange("A" + daterow + ":A" + daterow).setValue(y);
      Spent.getRange("B" + daterow + ":B" + daterow).setValue(m);
    }

    var dates = Spent.getRange("A" + daterow + ":C" + daterow).getValues();
    if (new Date(dates[0][0], dates[0][1] - 1, dates[0][2]) <= Today) {
      for (var s = 0; s < sum[0].length; s++)
        try {
          sum[0][s] = sum[0][s].replace(/([A-Za-z]+)([0-9]+)/g, '$1' + daterow);
        } catch (e) { }

      Spent.getRange("F" + daterow + ":Z" + daterow).setFormulasR1C1(sum);
      StartSpent = daterow;
    }

    daterow--;
  }

  calc();

  await log('...filters');

  var z2formula = parseInt((Spent.getRange("Z2:Z2").getFormula() ?? "SUM(Z3:Z99)").split("Z")[2]).toString();
  var maxRow = Spent.getMaxRows() - 1;
  var rangeVal = document.getElementById("range").value;
  document.getElementById("range").innerHTML = (z2formula > 99 ? `
<option value="99" selected>Top 99</value>
  ` : '') + `
<option value="${z2formula}">ZRange (${z2formula})</value>
  ` + (z2formula < 999 ? `
<option value="999">Top 999</value>
  ` : '')
    + (maxRow > 999 ? `
<option value="${maxRow}">End (${maxRow})</value>
  ` : '');
  if (rangeVal)
    document.getElementById("range").value = rangeVal;

  var colVal = document.getElementById("cols").value;
  document.getElementById("cols").innerHTML = '<option value="-1" selected>All</value>';
  var topRowFilters = Spent.getRange("F1:Z2").getValues();
  for (let i = 0; i < topRowFilters[0].length; i++) {
    if (topRowFilters[0][i] && topRowFilters[0][i].length > 0) {
      let total = topRowFilters[1][i] ?? "";
      if (total.toFixed)
        total = total.toFixed(2);
      document.getElementById("cols").innerHTML += `
<option value="${i+5}">${topRowFilters[0][i]}${total ? ' ('+Pound+total+')' : ''}</value>
      `;
    }
  }
  if (colVal)
    document.getElementById("cols").value = colVal;

  var prefVal = document.getElementById("prefixes").value;
  document.getElementById("prefixes").innerHTML = '<option value="" selected>All</value>';
  var pref = Spent.getRange("E4:E" + maxRow).getValues();
  pref.map(s => s[0]).filter(s => ~s.indexOf(':')).map(s => s.split(':')[0]).filter((a,i,c) => c.indexOf(a) == i).sort().forEach(s => {
    document.getElementById("prefixes").innerHTML += `
<option value="${s}:">${s}:</value>
      `;
  });
  if (prefVal)
    document.getElementById("prefixes").value = prefVal;
}

async function setupOneYear() {
  await log('One year...');

  await log('...clearing');

  // blank over sheet and recalc
  var r = ((Year - Beginning.getFullYear() - 1) * 12) + 2;
  Over.getRange("H" + r + ":L" + (r + 11)).setFormula("=0");
  Over.getRange("L" + r + ":L" + r).setFormula("=0");

  await log('...calc');
  calc();

  await log('...future zero')
  futureYear(0);
  await log('...overspent');
  overSpent();
  await log('...forecasting');
  overYearly();
  await log('...calc');
  calc();
  await log('...done');
}

async function setupLimited() {
  await log('F');

  // calc any Fn sheets if they exist, ie F3 using plan items with values <= 3
  for (var priority = 1; priority <= 9; priority++) {
    Future = Workbook.Sheets["F" + priority];
    if (typeof Future == "undefined" || Future == null)
      continue;

    log(priority + "...");
    futureYear(1, priority);
    nowData(priority);
  }
  await timeout();

  Future = Workbook.Sheets['Future'];
  await log('...calc');
  calc();
  await log('...done');
}

async function setupAllYears() {
  await log('All years...');
  var len = parseInt(Spent.getRange("G1:G1").getValue() ?? "3");
  futureYear(len);
  await log('...overspent');
  overSpent();
  await log('...forecasting');
  overYearly();
  await log('...breakdown');
  nowData();
  await log('...calc');
  calc();
  await log('...done');
}

async function Saving() {
  await log('Saving...');

  Workbook.Props.Author = "HuwSy/Money-XLS";
  Workbook.Props.CreatedDate = new Date(2021, 2, 7, 8, 35);
  Workbook.Props.ModifiedDate = new Date();
  Workbook.Props.LastAuthor = "HuwSy/Money-XLS";

  // hiding
  if (Year > Beginning.getFullYear() + 1)
    Over.hideRows(2, ((Year - Beginning.getFullYear() - 1) * 12));
  var future = Future.getRange("A2:G" + + Future.getMaxRows()).getValues();
  var f = 0;
  while (f < future.length && new Date(future[f][0], future[f][1] - 1, future[f][2]) <= Today) {
    f++;
  }
  if (f > 0)
    Future.hideRows(2, 2 + (f > 10 ? f - 10 : 0));

  // formatting
  var currentCols = [];
  currentCols[0] = { width: 45 / 7 };
  currentCols[1] = { width: 25 / 7 };
  currentCols[2] = { width: 25 / 7 };
  currentCols[3] = { width: 100 / 7 };
  currentCols[4] = { width: 175 / 7 };
  currentCols[5] = { width: 100 / 7 };
  currentCols[6] = { width: 25 / 7 };
  var headings = Spent.getRange("H1:X1").getValues();
  for (var heading = 0; heading < 17; heading++)
    currentCols.push((headings[0][heading] || '') == '' ? { width: 25 / 7 } : { width: 85 / 7 });
  currentCols[24] = { width: 100 / 7 };
  currentCols[25] = { width: 100 / 7 };

  Spent['!cols'] = currentCols;
  Spent['!autofilter'] = { ref: "F1:Z1" };
  Spent.fixHeight(1, Spent.getMaxRows());

  Plan['!cols'] = [{ width: 25 / 7 }, { width: 25 / 7 }, { width: 25 / 7 }, { width: 45 / 7 }, { width: 25 / 7 }, { width: 25 / 7 }, { width: 45 / 7 }, { width: 25 / 7 }, { width: 25 / 7 }, { width: 100 / 7 }, { width: 175 / 7 }, { width: 25 / 7 }, { width: 125 / 7 }, { width: 125 / 7 }, { width: 100 / 7 }];
  Plan['!autofilter'] = { ref: "A1:O1" };
  Plan.fixHeight(1, Plan.getMaxRows());

  for (var priority = 9; priority >= 0; priority--) {
    Future = Workbook.Sheets[priority > 0 ? "F" + priority : "Future"];
    if (typeof Future == "undefined" || Future == null)
      continue;
    Future['!cols'] = [{ width: 45 / 7 }, { width: 25 / 7 }, { width: 25 / 7 }, { width: 100 / 7 }, { width: 175 / 7 }, { width: 100 / 7 }, { width: 100 / 7 }];
  }
  Over['!cols'] = [{ width: 45 / 7 }, { width: 25 / 7 }, { width: 100 / 7 }, { width: 100 / 7 }, { width: 100 / 7 }, { width: 100 / 7 }, { width: 25 / 7 }, { width: 100 / 7 }, { width: 100 / 7 }, { width: 100 / 7 }, { width: 60 / 7 }, { width: 60 / 7 }];

  Over['!merges'] = [];
  var len = parseInt(Spent.getRange("G1:G1").getValue() ?? "3") + 1;
  for (var col = 6; col <= 11; col++) {
    for (var y = 1; y <= len; y++) {
      Over['!merges'].push({
        s: {
          c: col,
          r: 1 + (y - 1) * 12
        },
        e: {
          c: col,
          r: y * 12
        }
      });
    }
  }

  Spent.getRange("B1:C1").setNumberFormat("00");
  Spent.getRange("D1:D1").setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
  Spent.getRange("E3:E3").setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");

  var mRows = Spent.getMaxRows();
  for (var m = 2; m <= mRows; m++) {
    try {
      Spent.getRange("B" + m + ":C" + m).setNumberFormat("00");
    } catch (e) { }
    try {
      Spent.getRange("D" + m + ":D" + m).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
    } catch (e) { }
    try {
      Spent.getRange("F" + m + ":Z" + m).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
    } catch (e) { }
  }

  Plan.getRange("A2:I" + Plan.getMaxRows()).setNumberFormat("00");
  mRows = Plan.getMaxRows();
  for (var m = 1; m <= mRows; m++) {
    try {
      Plan.getRange("J" + m + ":J" + m).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
    } catch (e) { }
    try {
      Plan.getRange("M" + m + ":N" + m).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
    } catch (e) { }
  }

  await log('...outputting');

  // saving buffer
  function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  // saving file
  var format = FileName.substring(FileName.lastIndexOf('.') + 1);
  var wb = XLSX.write(Workbook, { bookType: format, cellStyles: true, compression: true, bookSST: true, dense: true, type: 'binary' });
  var buf = s2ab(wb);
  var blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", charset: "UTF-8", name: FileName });

  // downloading
  var aElement = document.createElement('a');
  aElement.innerHTML = FileName;
  aElement.download = FileName;
  aElement.setAttribute('download', FileName);
  aElement.href = URL.createObjectURL(blob);
  aElement.target = "_blank";

  await log('...done');

  document.getElementById('output').appendChild(aElement);
  await log('');
  aElement.click();
}

async function fileLoaded(e) {
  // loading xlsx
  Workbook = null, Spent = null, Future = null, Over = null, Plan = null;
  Workbook = XLSX.read(e.target.result, { type: 'array', cellStyles: false });
  await log('...loaded ' + Workbook.SheetNames.length + ' sheets');

  // sheet functions to mimic gs
  Workbook.SheetNames.forEach(sheetConfig);

  // sheets
  Spent = Workbook.Sheets['Spent'];
  Future = Workbook.Sheets['Future'];
  Over = Workbook.Sheets['Over'];
  Plan = Workbook.Sheets['Plan'];

  SST = [];
  for (var s in Workbook.Strings) {
    if (Workbook.Strings[s] && Workbook.Strings[s].t && Workbook.Strings[s].t.length > 0)
      SST.push(Workbook.Strings[s].t);
  }
  SST = SST.sort((a, b) => a.length - b.length);

  var removeSum = Spent.getRange("D3:D3").getFormula() ?? "";
  var expanded = expandSumFormula(removeSum);
  Spent.getRange("D3:D3").setFormula(expanded);
  calc();

  document.getElementById('future').style.display = "block";
  document.getElementById('futures').style.display = "block";
  document.getElementById('over').style.display = "block";
  document.getElementById('chart').style.display = "block";
  document.getElementById('options').style.display = "block";

  await setupSpentFields();
  await Filter(true);

  await log('Ready...');
}

function expandSumFormula(formula) {
  // Match the pattern: optional sign, SUM, and its arguments
  const match = formula.match(/([-+]*)SUM\(([^)]+)\)/i);
  if (!match) return formula;
  
  const sign = match[1] || '';
  const args = match[2];
  
  // Split by commas to get individual arguments
  const parts = args.split(',').map(p => p.trim());
  const expanded = [];
  
  parts.forEach(part => {
    // Check if it's a range (e.g., D4:D9)
    if (part.includes(':')) {
      const [start, end] = part.split(':');
      const colStart = start.match(/[A-Z]+/)[0];
      const rowStart = parseInt(start.match(/\d+/)[0]);
      const rowEnd = parseInt(end.match(/\d+/)[0]);
      
      // Expand the range
      for (let i = rowStart; i <= rowEnd; i++) {
        expanded.push(colStart + i);
      }
    } else {
      // Single cell reference
      expanded.push(part);
    }
  });
  
  // Join with the sign
  return sign + expanded.join(sign + (sign === '' ? '+' : ''));
}

function futureYear(len, priority) {
  // how much data we will store
  var EndDate = new Date(Year + len, Beginning.getMonth() + 1 - 1, 1);
  var NotBefore = new Date(Beginning.toJSON().substring(0, 10));
  if (len <= 0)
    NotBefore = new Date(Year - 1, Beginning.getMonth() + 1 - 1, 1);
  if (NotBefore < Beginning)
    NotBefore = Beginning;

  var future = [];
  var plan = Plan.getRange("A2:L" + Plan.getMaxRows()).getValues();

  // loop all planned items
  for (p = 0; p < plan.length; p++) {
    // if limited skip lower priority
    if ((priority == null && plan[p][11] == "-1")
      || (priority != null && plan[p][11] > priority))
      continue;

    // if there is start date and details
    if (plan[p][6] != "" && plan[p][7] != "" && plan[p][8] != "" && plan[p][10] != "") {
      var t = new Date(plan[p][6], plan[p][7] - 1, plan[p][8]);
      if (t >= EndDate) continue;
      var c = plan[p][11] == "-1" ? "" : plan[p][9];
      var n = plan[p][11] == "-1" ? "" : plan[p][10];

      var ay = plan[p][0];
      if (ay == '') ay = 0;
      var am = plan[p][1];
      if (am == '') am = 0;
      var ad = plan[p][2];
      if (ad == '') ad = 0;

      var e = EndDate;

      // get end date if exists
      if (plan[p][3] != "" && plan[p][4] != "" && plan[p][5] != "") {
        e = new Date(plan[p][3], plan[p][4] - 1, plan[p][5]);
      }

      // if no repeat use current date
      if (ay + am + ad <= 0) {
        e = t;
      }

      // override end date with EndDate
      if (e > EndDate) {
        e = EndDate;
      }

      // speed up one year calc or shorrter beginning dates
      if (NotBefore != null) {
        if (e < NotBefore || (e.toUTCString() == NotBefore.toUTCString() && ay + am + ad > 0))
          continue;

        if (ay + am + ad > 0)
          while (t < NotBefore) {
            t = new Date(t.getFullYear() + ay, t.getMonth() + am, t.getDate() + ad);
          }

        if (t >= EndDate)
          continue;
      }

      if (t > e || (t.toUTCString() == e.toUTCString() && ay + am + ad > 0))
        continue;

      // populate future and loop until end date
      do {
        var f = future.length;
        future[f] = [];
        future[f][0] = t.getFullYear();
        future[f][1] = t.getMonth() + 1;
        future[f][2] = t.getDate();
        future[f][3] = c;
        future[f][4] = n;

        t = new Date(t.getFullYear() + ay, t.getMonth() + am, t.getDate() + ad);
      } while (t < e);
    }
  }

  future.sort(function (a, b) {
    if (a[0] == b[0] && a[1] == b[1])
      return a[2] - b[2];
    if (a[0] == b[0])
      return a[1] - b[1];
    return a[0] - b[0];
  });

  Future.showRows(2, Future.getMaxRows());
  Future.getRange("A2:G" + Future.getMaxRows()).clear();
  Future.getRange("A2:E" + (future.length + 1)).setValues(future);

  Future.getRange("A2:C" + Future.getMaxRows()).setNumberFormat("00");
  Future.getRange("D2:D" + Future.getMaxRows()).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
}

function overSpent() {
  var spent = Spent.getRange("A4:F" + Spent.getMaxRows()).getValues();
  var future = Future.getRange("A2:E" + Future.getMaxRows()).getValues();

  var over = [];
  // loop future output
  for (var f = 0; f < future.length;) {
    var cc = future[f++];
    // if data exists
    if (cc[0] != "" && cc[1] != "" && cc[2] != "" && new Date(cc[0], cc[1] - 1, cc[2]) <= Today) {
      // create over row
      var o = over.length;
      over[o] = [];
      over[o][0] = cc[0];
      over[o][1] = cc[1];
      over[o][2] = 0.0; // spent scheet col d
      over[o][3] = 0.0; // spent f
      over[o][4] = parseFloat(cc[3] || 0);
      over[o][5] = 0.0; // [2] - [3] - [4]

      // loop future summing into [4]
      while (f < future.length && future[f][0] == cc[0] && future[f][1] == cc[1] && new Date(future[f][0], future[f][1] - 1, future[f][2]) <= Today) {
        over[o][4] += parseFloat(future[f++][3] || 0);
      }

      // find end of spent for this over row
      var s = spent.length - 1;
      while (s >= 0 && (spent[s][0] != cc[0] || spent[s][1] != cc[1])) s--;

      // loop spent into [2,3]
      while (s >= 0 && spent[s][0] == cc[0] && spent[s][1] == cc[1] && new Date(spent[s][0], spent[s][1] - 1, spent[s][2]) <= Today) {
        over[o][2] += parseFloat(spent[s][3] || 0);
        over[o][3] += parseFloat(spent[s][5] || 0);
        s--;
      }

      // calculate [5]
      over[o][5] = over[o][2] - over[o][3] - over[o][4];
    }
  }

  Over.getRange("A2:F" + Over.getMaxRows()).clear();
  Over.getRange("I2:I" + Over.getMaxRows()).clear();

  var start = (((over[0][0] - Beginning.getFullYear()) * 12) + 2);
  Over.getRange("A" + start + ":F" + (start + over.length - 1)).setValues(over);

  Over.getRange("A" + start + ":B" + (start + over.length - 1)).setNumberFormat("00");
  Over.getRange("C" + start + ":F" + (start + over.length - 1)).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
}

function overYearly() {
  var future = Future.getRange("A2:E" + Future.getMaxRows()).getValues();
  var over = [];
  // loop future output
  for (var f = 0; f < future.length && future[f][1] != "";) {
    var o = over.length;
    over[o] = [];
    over[o][0] = parseFloat(future[f][0] || 0) + 1;
    over[o][1] = 0.0;

    // while its same year sum up
    while (f < future.length && future[f][0] != "" && future[f][0] < over[o][0]) {
      over[o][1] += parseFloat(future[f][3] || 0);
      f++;
    }

    // while its before month sum up
    while (f < future.length && future[f][0] != "" && future[f][1] != "" && future[f][1] < Beginning.getMonth() + 1 && future[f][0] == over[o][0]) {
      over[o][1] += parseFloat(future[f][3] || 0);
      f++;
    }

    // loop back to next year
  }

  if (Year > Beginning.getFullYear() + 1)
    Over.showRows(2, ((Year - Beginning.getFullYear()) * 12) + 1);

  // output merged cells, formulas etc
  for (var i = 0; i < over.length; i++) {
    if (i == over.length - 1 && over[i][1] == 0)
      continue;
    var s = (((over[i][0] - Beginning.getFullYear() - 1) * 12) + 2);
    var e = s + 11;

    Over.getRange("A" + s + ":A" + s).setValue((over[i][0] - 1));
    Over.getRange("A" + e + ":A" + e).setValue(over[i][0] - (Beginning.getMonth() > 0 ? 0 : 1));
    Over.getRange("H" + s + ":H" + s).setFormula("=SUM(F" + s + ":F" + e + ")");
    Over.getRange("H" + s + ":H" + s).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
    Over.getRange("I" + s + ":I" + s).setValue(over[i][1]);
    Over.getRange("I" + s + ":I" + s).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
    Over.getRange("K" + s + ":K" + s).setFormula("=IF(TODAY()-DATE(A" + s + "," + (Beginning.getMonth() + 1) + ",1)<1,52,IF(TODAY()-DATE(A" + s + "," + (Beginning.getMonth() + 1) + ",1)>366,52,(TODAY()-DATE(A" + s + "," + (Beginning.getMonth() + 1) + ",1))/7))");
    Over.getRange("K" + s + ":K" + s).setNumberFormat('00');
    Over.getRange("L" + s + ":L" + s).setFormula("=H" + s + "/K" + s + "");
    Over.getRange("L" + s + ":L" + s).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
  }
}

function nowData(priority) {
  var future = Future.getRange("A2:G" + + Future.getMaxRows()).getValues();

  // find today and before beginning
  var f = 0, april = 1;
  while (f < future.length && new Date(future[f][0], future[f][1] - 1, future[f][2]) <= Today) {
    f++;
    if (f >= future.length || future[f][0] < (Year - 1) || (future[f][1] < Beginning.getMonth() + 1 && future[f][0] == (Year - 1)))
      april++;
  }

  var nfuture = [];

  // find last 10 items incase they haven passed yet
  var last10 = (f - 10) > 0 ? (f - 10) : 0;
  var x = last10 < april ? last10 : april;
  for (var p = x; p < f; p++) {
    nfuture[nfuture.length] = future[p];
  }

  // current balance
  var cur = 0.0, blanks = 4;
  if (priority == null) {
    // calc from spent
    cur = Spent.getRange("D1:D1").getValue() + Spent.getRange("D2:D2").getValue() - Spent.getRange("D3:D3").getValue();
    while (Spent.getRange("A" + blanks + ":A" + blanks).getValue() == "" && blanks < Spent.getMaxRows()) {
      cur -= parseFloat(Spent.getRange("D" + blanks + ":D" + blanks).getValue() || 0);
      blanks++;
    }
  } else {
    // calc from current month only
    blanks = nfuture.length - 1;
    while (blanks >= 0 && nfuture[blanks][4] != "") {
      cur += parseFloat(nfuture[blanks][3] || 0);
      blanks--;
    }
  }

  var minim = [], mincr = cur, minyr = Year;
  //today.setHours(0, 0, 0, 0);
  nfuture[nfuture.length] = [Today.getFullYear(), Today.getMonth() + 1, Today.getDate(), '', '', Today, cur];

  // find all future balance based off current
  for (var p = f; p < future.length && future[p][0] != ""; p++) {
    var n = nfuture.length;
    nfuture[n] = future[p];
    nfuture[n][5] = new Date(nfuture[n][0], nfuture[n][1] - 1, nfuture[n][2]);
    if (nfuture[n][4] == "") {
      if (priority != null) {
        cur = 0.0;
        nfuture[n][5] = "";
      }
    } else
      cur += parseFloat(nfuture[n][3] || 0);
    nfuture[n][6] = cur;
    if (nfuture[n][0] < minyr || (nfuture[n][1] < Beginning.getMonth() + 1 && nfuture[n][0] == minyr)) {
      if (cur < mincr)
        mincr = cur;
    } else {
      minim.push([mincr, minyr]);
      mincr = cur;
      minyr = nfuture[n][0] + 1;
    }
  }
  minim.push([mincr, minyr]);

  // output low points to over if not in Fn sheet
  if (priority == null) {
    Over.getRange("J2:J" + Over.getMaxRows()).clear();

    var r = ((Year - Beginning.getFullYear() - 1) * 12) + 2;
    for (var m in minim) {
      Over.getRange("A" + r + ":A" + r).setValue((minim[m][1] - 1));
      Over.getRange("A" + (r + 11) + ":A" + (r + 11)).setValue(minim[m][1] - (Beginning.getMonth() > 0 ? 0 : 1));

      Over.getRange("J" + r + ":J" + r).setValue(minim[m][0]);
      Over.getRange("J" + r + ":J" + r).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
      r += 12;
    }
  }

  Future.getRange("A2:G" + Future.getMaxRows()).clear();
  Future.getRange("A2:G" + (nfuture.length + 1)).setValues(nfuture);

  Future.getRange("A2:C" + (nfuture.length + 1)).setNumberFormat("00");
  Future.getRange("D2:D" + (nfuture.length + 1)).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
  Future.getRange("G2:G" + (nfuture.length + 1)).setNumberFormat(Pound + "#,##0.00;[Red]-" + Pound + "#,##0.00");
  Future.getRange("F2:F" + (nfuture.length + 1)).setNumberFormat("yyyy-mm-dd");

  if (f - x - 10 >= 2)
    Future.hideRows(2, f - x - 10);
}

function timeout(ms) {
  return new Promise(resolve => setTimeout(resolve, ms || 10));
}

function calc() {
  XLSX_CALC(Workbook);
}

function sheetConfig(sheetName) {
  // Mimics some of google sheets api that this was originally developed against
  var s = Workbook.Sheets[sheetName];

  s.getMaxRows = function () {
    try {
      return parseInt(s['!ref'].match(/[0-9]*$/)[0]);
    } catch (e) {
      return 64000;
    }
  }

  s.getRange = function (r) {
    r = r.match(/([A-Z]+)([0-9]+):*([A-Z]*)([0-9]*)/);
    r = [r[1], parseInt(r[2]), r[3] || r[1], parseInt(r[4] || r[2])];
    var r0 = r[0];

    function nextChar(c) {
      return XLSX.utils.encode_col(XLSX.utils.decode_col(c) + 1);
    }

    function getter(field) {
      var ret = [];
      while (r[3] >= r[1]) {
        var p = [];
        while (('0000' + r[2]).slice(-5) >= ('0000' + r[0]).slice(-5)) {
          try {
            var o = s[r[0] + r[1]] || {};
            if (o.t == 'e')
              p.push(null);
            else if (o[field])
              p.push(o[field]);
            else
              p.push(field == 'f' ? null : o.t == 'n' ? 0.0 : o.t == 'd' ? null : '');
          } catch (e) {
            p.push(null);
          }
          r[0] = nextChar(r[0]);
        }
        ret.push(p);
        r[1]++;
        r[0] = r0;
      }
      return ret;
    }

    function setter(field, v) {
      if (s.getMaxRows() < r[3])
        s['!ref'] = s['!ref'].replace(/[0-9]*$/, r[3]);

      var a = 0, b, formula_ref = {}, cells = [];
      while (r[3] >= r[1]) {
        b = 0;
        while (('0000' + r[2]).slice(-5) >= ('0000' + r[0]).slice(-5)) {
          var o = r[0] + r[1]
          try {
            if (!s[o])
              s[o] = {};

            if (field == 'c') {
              var formula = formula_ref[sheetName + '!' + o] = {
                formula_ref: formula_ref,
                wb: Workbook,
                sheet: s,
                sheet_name: sheetName,
                cell: s[o],
                name: o,
                status: 'new',
                exec_formula: (v || {}).exec_formula
              };
              cells.push(formula);
            } else if (field == 'w') {
              s[o].z = v;
              if (s[o].v && v == '00')
                s[o].w = (s[o].v < 10 ? '0' : '') + s[o].v;
              else if (s[o].v && v == 'yyyy-mm-dd')
                s[o].w = (new Date(s[o].v)).toJSON().substring(0, 10);
              else if (s[o].v && v.indexOf(Pound) >= 0)
                s[o].w = (s[o].v < 0 ? "-" : "") + Pound + parseFloat(s[o].v || 0).toFixed(2);
              else
                s[o].w = XLSX.SSF.format(v, s[o].v)
            } else if (field == null || !v[a] || !v[a][b]) {
              delete s[o].v;
              delete s[o].f;
              delete s[o].w;
            } else if (field == 'v') {
              s[o].t = typeof v[a][b] == 'number' ? 'n' : typeof v[a][b] == 'object' ? 'd' : 't';
              s[o].v = v[a][b];
              s[o].w = v[a][b].toString();
            } else if (field == 'f') {
              // TODO: this is a dirty hack for what is needed specifically not actual fix
              s[o].f = v[a][b] ? v[a][b]
                .replace(/(\$)([A-Z]+)3/g, '$1$2' + r[1])
                .replace(/^=/, '')
                : null;
              s[o].t = v[a][b] ? 's' : 't';
              s[o].v = '';
              s[o].w = '';
            }
          } catch (e) { }
          r[0] = nextChar(r[0]);
          b++;
        }
        r[1]++;
        r[0] = r0;
        a++;
      }
      return cells;
    }

    return {
      getValue: function () {
        return getter('v')[0][0];
      },
      setValue: function (v) {
        setter('v', [[v]]);
      },
      getValues: function () {
        return getter('v');
      },
      setValues: function (v) {
        setter('v', v);
      },
      clear: function () {
        setter(null, null);
      },
      getFormula: function () {
        return getter('f')[0][0];
      },
      setFormula: function (v) {
        setter('f', [[v]]);
      },
      getFormulasR1C1: function () {
        return getter('f');
      },
      setFormulasR1C1: function (v) {
        setter('f', v);
      },
      setNumberFormat: function (f) {
        setter('w', f.replace(Pound, '"' + Pound + '"').replace(Pound, '"' + Pound + '"'));
      }
    }
  }

  function rowsSet(b, e, v, t) {
    s['!rows'] = s['!rows'] || {};
    b--;
    while (b <= e) {
      s['!rows'][b] = s['!rows'][b] || {};
      s['!rows'][b][t] = v;
      b++;
    }
  }

  s.hideRows = function (b, e) {
    rowsSet(b, e, true, 'hidden');
  }
  s.showRows = function (b, e) {
    rowsSet(b, e, false, 'hidden');
  }
  s.fixHeight = function (b, e) {
    rowsSet(b, e, 15, 'hpt');
  }
}

function Init() {
  document.getElementById('start').value = Beginning.toJSON().substring(0, 10);
  document.getElementById('today').value = Today.toJSON().substring(0, 10);

  Spending = document.getElementById("spending").getElementsByTagName('tr');
  Upcoming = document.getElementById("upcoming").getElementsByTagName('tr');
  Charts = document.getElementById("chart").getElementsByTagName('canvas');
  Plans = document.getElementById("plans").getElementsByTagName('tr');
  Formulas = document.getElementById("formulas").getElementsByTagName('div');

  XLSX_CALC.import_functions({
    'FIND': function (a, b, c) {
      return (b || '').toString().indexOf((a || '').toString(), parseInt(c || 0));
    },
    'LOWER': function (a) {
      return (a || '').toString().toLowerCase();
    },
    'ISERROR': function (a) {
      return a == -1;
    },
    'DATE': function (y, m, d) {
      return new Date(parseInt(y), parseInt(m) - 1, parseInt(d), 0, 0, 0, 0);
    },
    'NOT': function (a) {
      return !a;
    },
    'TODAY': function () {
      var t = new Date();
      t.setHours(0, 0, 0, 0);
      return t;
    }
  }, { override: true });
}

document.addEventListener("DOMContentLoaded", Init);
