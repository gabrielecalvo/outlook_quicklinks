const items = window.items;
document.getElementById("sideload-msg").style.display = "none";
document.getElementById("app-body").style.display = "flex";
fillList();

function fillList() {
  var ul = document.getElementById("quicklinkslist");

  items.forEach(function(obj) {
    if ("icon" in obj) {
      createSecionHeader(ul, obj);
    } else {
      createItemLink(ul, obj);
    }
  });
}

function createSecionHeader(ul, obj) {
  let p = document.createElement("p");
  let span = document.createElement("span");
  let i = document.createElement("i");
  i.classList.add("ms-Icon");
  i.classList.add("ms-Icon--" + obj.icon);
  i.classList.add("ms-font-xl");
  span.classList.add("ms-font-l");
  span.innerHTML += " " + obj.text;
  p.appendChild(i);
  p.appendChild(span);
  ul.appendChild(p);
}

function createItemLink(ul, obj) {
  let a = document.createElement("a");
  var subject = "subject" in obj ? obj.subject : obj.text;
  a.href = "mailto:" + obj.email + "?subject=" + subject;
  a.innerHTML += obj.text;

  let li = document.createElement("li");
  li.classList.add("ms-ListItem");
  li.appendChild(a);
  ul.appendChild(li);
}
