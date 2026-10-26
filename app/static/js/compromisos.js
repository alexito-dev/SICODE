(() => {
    const $ = (id) => document.getElementById(id);
    const CAMPOS = ["ruc", "razon_social", "telefono", "contacto", "fecha_llamada", "resultado",
        "fecha_compromiso", "monto_compromiso", "estado", "proxima_llamada", "observaciones"];
    let datos = [];

    const money = (n) => (n || 0).toLocaleString("es-PE", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
    const dmy = (iso) => (iso ? iso.slice(0, 10).split("-").reverse().join("/") : "");

    function params() {
        const p = new URLSearchParams();
        [["q", "fQ"], ["resultado", "fResultado"], ["estado", "fEstado"], ["desde", "fDesde"], ["hasta", "fHasta"]]
            .forEach(([k, id]) => { if ($(id).value) p.set(k, $(id).value); });
        return p.toString();
    }

    async function cargar() {
        const r = await fetch("/compromisos/api/llamadas?" + params());
        const j = await r.json();
        datos = j.data || [];
        const s = j.resumen || {};
        $("kLlamadas").textContent = s.llamadas;
        $("kCompromisos").textContent = s.compromisos;
        $("kTotal").textContent = money(s.monto_total);
        $("kPendiente").textContent = money(s.monto_pendiente);
        $("kVencido").textContent = money(s.monto_vencido);
        $("kCumplido").textContent = money(s.monto_cumplido);
        $("kHoy").textContent = s.vencen_hoy;
        $("kLlamar").textContent = s.por_llamar;
        pintar();
    }

    function pintar() {
        $("vacio").style.display = datos.length ? "none" : "block";
        $("tbody").innerHTML = datos.map((d) => `
            <tr>
                <td>${esc(d.fecha_llamada.replace("T", " ").slice(0, 16).split(" ").map((x, i) => i ? x : dmy(x)).join(" "))}</td>
                <td>${esc(d.ruc)}</td>
                <td>${esc(d.razon_social)}</td>
                <td>${esc(d.telefono)}${d.contacto ? "<br><small>" + esc(d.contacto) + "</small>" : ""}</td>
                <td>${esc(d.resultado)}</td>
                <td>${dmy(d.fecha_compromiso)}</td>
                <td class="num">${d.monto_compromiso === null || d.monto_compromiso === undefined ? "" : money(d.monto_compromiso)}</td>
                <td><span class="cp-tag ${esc(d.estado_efectivo)}">${esc(d.estado_efectivo)}</span></td>
                <td>${dmy(d.proxima_llamada)}</td>
                <td class="obs">${esc(d.observaciones)}</td>
                <td class="cp-fila-acc">
                    <button title="Editar" data-edit="${d.id}"><i class="fa-solid fa-pen"></i></button>
                    <button class="del" title="Eliminar" data-del="${d.id}"><i class="fa-solid fa-trash"></i></button>
                </td>
            </tr>`).join("");
    }

    function ahoraLocal() {
        const d = new Date();
        d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
        return d.toISOString().slice(0, 16);
    }

    function abrir(d) {
        $("modalTitulo").textContent = d ? "Editar llamada" : "Nueva llamada";
        $("id").value = d ? d.id : "";
        CAMPOS.forEach((c) => { $(c).value = d ? (d[c] ?? "") : ""; });
        if (!d) {
            $("fecha_llamada").value = ahoraLocal();
            $("resultado").selectedIndex = 0;
            $("estado").selectedIndex = 0;
        }
        $("formError").textContent = "";
        $("modal").classList.add("open");
        $("ruc").focus();
    }

    const cerrar = () => $("modal").classList.remove("open");

    async function guardar(e) {
        e.preventDefault();
        const body = {};
        CAMPOS.forEach((c) => { body[c] = $(c).value; });
        body.monto_compromiso = body.monto_compromiso === "" ? null : parseFloat(body.monto_compromiso);
        const id = $("id").value;
        const r = await fetch("/compromisos/api/llamadas" + (id ? "/" + id : ""), {
            method: id ? "PUT" : "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
        });
        if (!r.ok) {
            const j = await r.json().catch(() => ({}));
            $("formError").textContent = typeof j.detail === "string" ? j.detail : "Revise los datos ingresados.";
            return;
        }
        cerrar();
        cargar();
    }

    let t;
    const debounce = () => { clearTimeout(t); t = setTimeout(cargar, 250); };

    $("btnNuevo").onclick = () => abrir(null);
    $("btnCancelar").onclick = cerrar;
    $("form").onsubmit = guardar;
    $("modal").onclick = (e) => { if (e.target === $("modal")) cerrar(); };
    $("btnExcel").onclick = () => { window.location = "/compromisos/api/exportar?" + params(); };
    $("btnLimpiar").onclick = () => {
        ["fQ", "fResultado", "fEstado", "fDesde", "fHasta"].forEach((id) => { $(id).value = ""; });
        cargar();
    };
    $("fQ").oninput = debounce;
    ["fResultado", "fEstado", "fDesde", "fHasta"].forEach((id) => { $(id).onchange = cargar; });
    $("tbody").onclick = async (e) => {
        const b = e.target.closest("button");
        if (!b) return;
        if (b.dataset.edit) abrir(datos.find((d) => String(d.id) === b.dataset.edit));
        if (b.dataset.del && confirm("¿Eliminar este registro?")) {
            await fetch("/compromisos/api/llamadas/" + b.dataset.del, { method: "DELETE" });
            cargar();
        }
    };

    cargar();
})();
