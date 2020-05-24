const items = window.items;

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    fillList();
  }
});

export async function fillList() {
  var ul = document.getElementById("quicklinkslist");

  items.forEach(function(obj) {
    if ("icon" in obj) {
      createSecionHeader(ul, obj);
    } else {
      createItemLink(ul, obj);
    }
  });
}

export async function createSecionHeader(ul, obj) {
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

export async function createItemLink(ul, obj) {
  let a = document.createElement("a");

  if (Office.context.mailbox.item.displayReplyForm != undefined) {
    // read mode
    var subject = "subject" in obj ? obj.subject : obj.text;
    a.href = "mailto:" + obj.email + "?subject=" + subject;
  } else {
    // compose mode
    a.href = "#";
    a.onclick = () => {
      addRecipient(obj.email);
    };
  }
  a.innerHTML += obj.text;

  let li = document.createElement("li");
  li.classList.add("ms-ListItem");
  li.appendChild(a);
  ul.appendChild(li);
}

export async function addRecipient(newRecipient) {
  var item = Office.context.mailbox.item;

  var callback = function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      document.getElementById("log").innerHTML = "<h4>" + asyncResult.error.message + "</h4><br/>";
      write(asyncResult.error.message);
    } else {
      setInterval(() => {
        Office.context.ui.closeContainer();
      }, 800);
    }
  };
  item.to.addAsync([newRecipient], callback);
}
