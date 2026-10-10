/* =========================================================================
   قالب طلب الطعام — منطق التطبيق (Restaurant Ordering Template — App Logic)
   الأقسام:
  1. الحالة
   2. الإعدادات → الصفحة
   3. الأصناف: التحميل من إكسل، التحليل، عرض القائمة
  4. السلة: إضافة/حذف/كمية، العرض
   5. إتمام الطلب: النافذة، رفع إيصال المحفظة، الإرسال → كائن الطلب
  6. التهيئة / ربط الأحداث
   ========================================================================= */

/* ---------------------------- 1. الحالة ---------------------------- */

let state = {
  products: [],          // [{id, category, name, description, price, image, available}]
  categories: [],         // أسماء التصنيفات الفريدة، بترتيب أول ظهور
  activeCategory: "الكل",
  cart: {},
  orderSeq: 0

};



const money = (n) => `${Number(n || 0).toFixed(2)} ${APP_CONFIG.currencySymbol}`;
const formatDate = (iso) => new Date(iso).toLocaleString("ar-EG-u-nu-latn");

/* ---------------------------- 2. الإعدادات → الصفحة ---------------------------- */

function applyConfigToDom(){
  document.title = `${APP_CONFIG.restaurantName} — اطلب أونلاين`;
  document.querySelectorAll("[data-config]").forEach(el => {
    const key = el.getAttribute("data-config");
    if (key === "phoneHref"){ el.setAttribute("href", APP_CONFIG.phoneHref); return; }
    if (key === "emailHref"){ el.setAttribute("href", APP_CONFIG.emailHref); return; }
    if (key === "whatsappHref"){ el.setAttribute("href", APP_CONFIG.whatsappHref); return; }
    if (APP_CONFIG[key] !== undefined) el.textContent = APP_CONFIG[key];
  });
}

/* ---------------------------- 3. الأصناف ---------------------------- */

function normalizeRow(row, index){
  // يقبل أسماء أعمدة عربية أو إنجليزية، بأي شكل في الكتابة أو المسافات
  const get = (...names) => {
    for (const n of names){
      const key = Object.keys(row).find(k => k.trim().toLowerCase().replace(/[\s_]/g,"") === n);
      if (key !== undefined && row[key] !== "") return row[key];
    }
    return "";
  };
  const name = get("name", "item", "dish", "product", "الاسم", "الصنف", "اسمالصنف");
  const category = get("category", "cat", "الفئة", "التصنيف", "القسم") || "القائمة";
  const price = parseFloat(get("price", "cost", "السعر")) || 0;
  const description = get("description", "desc", "الوصف");
  const image = get("imageurl", "image", "img", "photo", "رابطالصورة", "الصورة");
  const colors = parseOptions(get("color", "colors", "colour", "colours", "اللون", "الألوان"), ["أسود", "أبيض", "أحمر", "أزرق"]);
  const sizes = parseOptions(get("size", "sizes", "المقاس", "المقاسات"), ["S", "M", "L", "XL"]);
  const availableRaw = get("available", "instock", "active", "متوفر", "الحالة");
  const unavailableWords = ["no","false","0","out","soldout","لا","غيرمتوفر","نفذ","نفدت","غيرمتاح"];
  const available = availableRaw === "" ? true : !unavailableWords.includes(String(availableRaw).trim().toLowerCase().replace(/\s/g,""));
  const id = get("id") || `${category}-${name}-${index}`.toLowerCase().replace(/[^a-z0-9\u0600-\u06FF]+/g,"-");
  if (!name) return null;
  return { id, category, name, description, price, image, available, colors, sizes };
}

function workbookToProducts(workbook){
  const sheetName = workbook.SheetNames.find(n => ["products","الأصناف","المنتجات"].includes(n.trim().toLowerCase())) || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  return rows.map(normalizeRow).filter(Boolean);
}



async function loadProducts(){
  // تحميل ملف Google Sheets المنشور بصيغة CSV.
  try{
    const res = await fetch(APP_CONFIG.productsFileUrl, { cache: "no-store" });
    if (!res.ok) throw new Error("fetch failed");
    const csv = await res.text();
    const workbook = XLSX.read(csv, { type: "string" });
    const products = workbookToProducts(workbook);
    if (!products.length) throw new Error("empty sheet");
    applyProducts(products);

    return;
  }catch(err){


  }
}

function applyProducts(products){
  state.products = products;
  const seen = new Set();
  state.categories = [];
  products.forEach(p => { if (!seen.has(p.category)){ seen.add(p.category); state.categories.push(p.category); } });
  state.activeCategory = "الكل";
  renderCategories();
  renderMenu();
}

function renderCategories(){
  const wrap = document.getElementById("categoryPills");
  const cats = ["الكل", ...state.categories];
  wrap.innerHTML = cats.map(c =>
    `<button type="button" class="cat-pill ${c === state.activeCategory ? "active" : ""}" data-cat="${escapeAttr(c)}">${escapeHtml(c)}</button>`
  ).join("");
  wrap.querySelectorAll(".cat-pill").forEach(btn => {
    btn.addEventListener("click", () => {
      state.activeCategory = btn.getAttribute("data-cat");
      renderCategories();
      renderMenu();
    });
  });
}

function renderMenu(){
  const container = document.getElementById("menuSections");
  const emptyEl = document.getElementById("menuEmpty");
  const cats = state.activeCategory === "الكل" ? state.categories : [state.activeCategory];
  const visibleCats = cats.filter(c => state.products.some(p => p.category === c));

  if (!state.products.length){
    container.innerHTML = "";
    emptyEl.classList.remove("d-none");
    return;
  }
  emptyEl.classList.add("d-none");

  container.innerHTML = visibleCats.map(cat => {
    const items = state.products.filter(p => p.category === cat);
    return `
      <section class="menu-section" id="cat-${escapeAttr(slugify(cat))}">
        <h2 class="menu-section-title">${escapeHtml(cat)} <span class="count" dir="ltr">${items.length}</span></h2>
        <div class="menu-grid">
          ${items.map(dishCardHtml).join("")}
        </div>
      </section>`;
  }).join("");

  container.querySelectorAll("[data-add]").forEach(btn => {
    btn.addEventListener("click", () => addSelectedProductToCart(btn, "data-add"));
  });
  container.querySelectorAll("[data-inc]").forEach(btn => {
    btn.addEventListener("click", () => addSelectedProductToCart(btn, "data-inc"));
  });
  container.querySelectorAll("[data-dec]").forEach(btn => {
    btn.addEventListener("click", () => decrementCartItem(btn.getAttribute("data-dec")));
  });
}

function dishCardHtml(p){
  const qty = cartEntries().filter(e => e.product.id === p.id).reduce((sum, e) => sum + e.qty, 0);
  const img = p.image || "https://placehold.co/400x300?text=%20";
  return `
    <article class="dish-card ${p.available ? "" : "dish-unavailable"}">
      <img class="dish-img" src="${escapeAttr(img)}" alt="${escapeAttr(p.name)}" loading="lazy"
           onerror="this.src='https://placehold.co/400x300?text=%20'">
      <div class="dish-body">
        <h3 class="dish-name">${escapeHtml(p.name)}</h3>
        ${p.description ? `<p class="dish-desc">${escapeHtml(p.description)}</p>` : ""}
        <div class="variant-fields">
          <label>اللون
            <select class="form-select form-select-sm" data-color-for="${escapeAttr(p.id)}">
              ${p.colors.map(color => `<option value="${escapeAttr(color)}">${escapeHtml(color)}</option>`).join("")}
            </select>
          </label>
          <label>المقاس
            <select class="form-select form-select-sm" data-size-for="${escapeAttr(p.id)}">
              ${p.sizes.map(size => `<option value="${escapeAttr(size)}">${escapeHtml(size)}</option>`).join("")}
            </select>
          </label>
        </div>
        <div class="dish-footer">
          <span class="dish-price mono" dir="ltr">${money(p.price)}</span>
          ${p.available
            ? (qty > 0
                ? `<div class="qty-stepper">
                     <button type="button" data-dec="${p.id}" aria-label="إنقاص واحد">−</button>
                     <span class="mono" dir="ltr">${qty}</span>
                     <button type="button" data-inc="${p.id}" aria-label="إضافة واحد">+</button>
                   </div>`
                : `<button type="button" class="add-btn" data-add="${p.id}" aria-label="إضافة ${escapeAttr(p.name)} إلى السلة"><i class="bi bi-plus-lg"></i></button>`)
            : `<span class="badge-sold-out">غير متوفر</span>`}
        </div>
      </div>
    </article>`;
}

/* ---------------------------- 4. السلة ---------------------------- */

function addToCart(id, color, size){
  const key = cartKey(id, color, size);
  state.cart[key] = (state.cart[key] || 0) + 1;
  renderMenu();
  renderCart();
}
function addSelectedProductToCart(button, attribute){
  const card = button.closest(".dish-card");
  const id = button.getAttribute(attribute);
  addToCart(
    id,
    card.querySelector(`[data-color-for="${escapeAttr(id)}"]`).value,
    card.querySelector(`[data-size-for="${escapeAttr(id)}"]`).value
  );
}
function decrementCartItem(id){
  const entry = cartEntries().find(e => e.product.id === id);
  if (!entry) return;
  state.cart[entry.key] -= 1;
  if (state.cart[entry.key] <= 0) delete state.cart[entry.key];
  renderMenu();
  renderCart();
}
function removeFromCart(id){
  delete state.cart[id];
  renderMenu();
  renderCart();
}

function cartEntries(){
  return Object.entries(state.cart)
    .map(([key, qty]) => {
      const [id, color = "", size = ""] = key.split("::").map(decodeURIComponent);
      return { key, product: state.products.find(p => p.id === id), qty, color, size };
    })
    .filter(e => e.product);
}
function cartCount(){ return Object.values(state.cart).reduce((a,b) => a+b, 0); }
function cartSubtotal(){ return cartEntries().reduce((sum,e) => sum + e.product.price * e.qty, 0); }

function renderCart(){
  const count = cartCount();
  document.getElementById("cartCount").textContent = count;
  document.getElementById("cartFabCount").textContent = count;

  const entries = cartEntries();
  const itemsEl = document.getElementById("cartItems");
  const emptyEl = document.getElementById("cartEmptyMsg");
  const checkoutBtn = document.getElementById("checkoutBtn");

  if (!entries.length){
    itemsEl.innerHTML = "";
    emptyEl.classList.remove("d-none");
    checkoutBtn.disabled = true;
  }else{
    emptyEl.classList.add("d-none");
    checkoutBtn.disabled = false;
    itemsEl.innerHTML = entries.map(e => `
      <div class="cart-item">
        <img src="${escapeAttr(e.product.image || 'https://placehold.co/100x100?text=%20')}" alt="" onerror="this.src='https://placehold.co/100x100?text=%20'">
        <div>
          <div class="cart-item-name">${escapeHtml(e.product.name)} × <span dir="ltr">${e.qty}</span></div>
          <div class="cart-item-price" dir="ltr">${money(e.product.price)} / للقطعة</div>
          <div class="cart-item-variants">اللون: ${escapeHtml(e.color)} · المقاس: ${escapeHtml(e.size)}</div>
        </div>
        <button type="button" class="cart-item-remove" data-remove="${escapeAttr(e.key)}"><i class="bi bi-trash"></i></button>
      </div>`).join("");
    itemsEl.querySelectorAll("[data-remove]").forEach(btn => {
      btn.addEventListener("click", () => removeFromCart(btn.getAttribute("data-remove")));
    });
  }

  const subtotal = cartSubtotal();
  document.getElementById("cartSubtotal").textContent = money(subtotal);
  document.getElementById("cartTotal").textContent = money(subtotal);
}

/* ---------------------------- 5. إتمام الطلب ---------------------------- */

function openCheckout(){
  const entries = cartEntries();
  if (!entries.length) return;
  document.getElementById("reviewCount").textContent = cartCount();
  document.getElementById("reviewSubtotal").textContent = money(cartSubtotal());
  document.getElementById("reviewTotal").textContent = money(cartSubtotal());
  bootstrap.Offcanvas.getInstance(document.getElementById("cartPanel"))?.hide();
  new bootstrap.Modal(document.getElementById("checkoutModal")).show();
}



function nextOrderNumber(){
  state.orderSeq += 1;
  return state.orderSeq;
}

async function sendOrderToSheet(order){
  if (!APP_CONFIG.orderSubmitUrl) return true;
  try{
    await fetch(APP_CONFIG.orderSubmitUrl, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({ ...order, receiptAttached: Boolean(order.receiptDataUrl) })
    });
    return true;
  }catch(err){
    alert("تعذّر إرسال الطلب إلى المطعم. حاول مرة أخرى.");
    return false;
  }
}
function sendOrderToWhatsApp(order) {
  let message = `*طلب جديد 🛒*\n\n`;
   message += `*رقم الطلب:* ${order.id}\n`;
  message += `*الاسم:* ${order.customerName}\n`;
  message += `*الهاتف:* ${order.phone}\n`;
  
    message += `*العنوان:* ${order.address}\n`;
  

  if (order.notes) {
    message += `*الملاحظات:* ${order.notes}\n`;
  }

  message += `*الدفع:* ${order.paymentMethod}\n\n`;
  message += `*الأصناف:*\n`;

  order.items.forEach(item => {
    message += `- ${item.qty}x ${item.name} | اللون: ${item.color} | المقاس: ${item.size} (${item.price} ج.م)\n`;
  });

  message += `\n*المجموع الإجمالي:* ${order.total} ج.م`;

  // Encode the message to make it URL-safe
  const encodedMessage = encodeURIComponent(message);
  
  // Open WhatsApp in a new tab using the URL from config.js
  window.open(`${APP_CONFIG.whatsappHref}?text=${encodedMessage}`, '_blank');
}
async function handleCheckoutSubmit(e){
  e.preventDefault();
  const entries = cartEntries();
  if (!entries.length) return;

  const paymentMethod = document.getElementById("paymentMethod").value;
  const electronicPayment = ["Vodafone", "Fawry"].includes(paymentMethod);
 

  const orderNumber = nextOrderNumber();
  const order = {
    orderNumber,
    id: Math.random().toString(36).substring(2, 10).toUpperCase(),
    timestamp: new Date().toLocaleString('ar-EG'),
    customerName: document.getElementById("custName").value.trim(),
    phone: document.getElementById("custPhone").value.trim(),
    address: document.getElementById("custAddress").value.trim(),
    notes: document.getElementById("custNotes").value.trim(),
    paymentMethod,
    receiptDataUrl: electronicPayment ? state.pendingReceiptDataUrl : "",
    items: entries.map(e => ({ name: e.product.name, category: e.product.category, color: e.color, size: e.size, qty: e.qty, price: e.product.price, lineTotal: +(e.product.price * e.qty).toFixed(2) })),
    subtotal: +cartSubtotal().toFixed(2),
    total: +cartSubtotal().toFixed(2),
    status: "جديد"
  };

  const sheetPromise = sendOrderToSheet(order);
  sendOrderToWhatsApp(order);
  if (!await sheetPromise) return;

  // إعادة تعيين السلة والنموذج
  state.cart = {};
  renderMenu();
  renderCart();
 
  document.getElementById("checkoutForm").reset();
  document.getElementById("walletFields").classList.add("d-none");

  bootstrap.Modal.getInstance(document.getElementById("checkoutModal"))?.hide();
   new bootstrap.Modal(document.getElementById("confirmModal")).show();
}

/* ---------------------------- أدوات مساعدة ---------------------------- */

function escapeHtml(str=""){
  return String(str).replace(/[&<>"']/g, m => ({ "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;" }[m]));
}
function parseOptions(value, fallback){
  const options = String(value || "").split(/[,|؛\n]+/).map(option => option.trim()).filter(Boolean);
  return [...new Set(options.length ? options : fallback)];
}
function cartKey(id, color, size){
  return [id, color, size].map(encodeURIComponent).join("::");
}
function escapeAttr(str=""){ return escapeHtml(str); }
function slugify(str=""){ return str.toLowerCase().trim().replace(/[^a-z0-9\u0600-\u06FF]+/g,"-"); }

/* ---------------------------- 7. التهيئة ---------------------------- */

function bindEvents(){
  document.getElementById("cartOpenBtn").addEventListener("click", () => new bootstrap.Offcanvas(document.getElementById("cartPanel")).show());
  document.getElementById("cartFabBtn").addEventListener("click", () => new bootstrap.Offcanvas(document.getElementById("cartPanel")).show());
  document.getElementById("checkoutBtn").addEventListener("click", openCheckout);
  document.getElementById("checkoutForm").addEventListener("submit", handleCheckoutSubmit);
 
  document.getElementById("paymentMethod").addEventListener("change", (e) => {
    const electronicPayment = ["Vodafone", "Fawry"].includes(e.target.value);
    document.getElementById("walletFields").classList.toggle("d-none", !electronicPayment);
      });

 

}

function init(){
  applyConfigToDom();
  bindEvents();
  renderCart();
  loadProducts();
}

document.addEventListener("DOMContentLoaded", init);
