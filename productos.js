(function () {
  'use strict';

  const BLOB_BASE = 'https://nutriheal.blob.core.windows.net/productos/';
  const WHATSAPP_NUMERO = '573147088080';
  const MAX_VARIANTES = 5; // revisa _01 .. _05

  let catalogo = { categorias: [], productos: [] };
  let filtroActivo = { tipo: 'todos', categoriaId: null, subcategoriaId: null };

  const gridEl = document.getElementById('gridProductos');
  const statusEl = document.getElementById('gridStatus');
  const listaCategoriasEl = document.getElementById('listaCategorias');

  function formatPrecio(valor) {
    const n = Number(valor) || 0;
    return '$ ' + n.toLocaleString('es-CO', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function imagenPrincipalUrl(producto) {
    return BLOB_BASE +
      encodeURIComponent(producto.categoria) + '/' +
      encodeURIComponent(producto.subcategoria) + '/' +
      encodeURIComponent(producto.codigo) + '.png';
  }

  function imagenVarianteUrl(producto, n) {
    const sufijo = String(n).padStart(2, '0');
    return BLOB_BASE +
      encodeURIComponent(producto.categoria) + '/' +
      encodeURIComponent(producto.subcategoria) + '/' +
      encodeURIComponent(producto.codigo + '_' + sufijo) + '.png';
  }

  function whatsappUrl(producto) {
    const texto = `Hola, deseo adquirir el producto ${producto.nombre} (Código: ${producto.codigo})`;
    return `https://wa.me/${WHATSAPP_NUMERO}?text=${encodeURIComponent(texto)}`;
  }

  function checkImagenExiste(url) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload = () => resolve(true);
      img.onerror = () => resolve(false);
      img.src = url;
    });
  }

  /* ══════════════════════════════════════
     CARGA DE DATOS
     ══════════════════════════════════════ */
  fetch('catalogo.json')
    .then((r) => {
      if (!r.ok) throw new Error('No se pudo cargar catalogo.json');
      return r.json();
    })
    .then((data) => {
      catalogo = data;
      renderFiltros();
      renderGrid();
    })
    .catch((err) => {
      statusEl.textContent = 'No se pudo cargar el catálogo. Intenta de nuevo más tarde.';
      console.error(err);
    });

  /* ══════════════════════════════════════
     FILTROS (acordeón categoría → subcategoría)
     ══════════════════════════════════════ */
  function renderFiltros() {
    listaCategoriasEl.innerHTML = '';

    catalogo.categorias.forEach((cat) => {
      const item = document.createElement('div');
      item.className = 'cat-item';

      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = 'cat-toggle';
      toggle.innerHTML = `<span>${cat.nombre}</span><span class="chevron">▶</span>`;
      toggle.addEventListener('click', () => {
        item.classList.toggle('open');
      });

      const subList = document.createElement('ul');
      subList.className = 'subcat-list';

      cat.subcategorias.forEach((sub) => {
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = sub.nombre;
        btn.dataset.subcategoriaId = sub.id;
        btn.addEventListener('click', () => {
          filtroActivo = { tipo: 'subcategoria', categoriaId: cat.id, subcategoriaId: sub.id };
          marcarSubcategoriaActiva(sub.id);
          renderGrid();
        });
        li.appendChild(btn);
        subList.appendChild(li);
      });

      item.appendChild(toggle);
      item.appendChild(subList);
      listaCategoriasEl.appendChild(item);
    });
  }

  function marcarSubcategoriaActiva(subcategoriaId) {
    document.querySelectorAll('.subcat-list button').forEach((b) => {
      b.classList.toggle('active', Number(b.dataset.subcategoriaId) === Number(subcategoriaId));
    });
  }

  document.getElementById('btnLimpiarFiltro').addEventListener('click', () => {
    filtroActivo = { tipo: 'todos', categoriaId: null, subcategoriaId: null };
    document.querySelectorAll('.subcat-list button.active').forEach((b) => b.classList.remove('active'));
    renderGrid();
  });

  /* ══════════════════════════════════════
     GRID DE PRODUCTOS
     ══════════════════════════════════════ */
  function productosFiltrados() {
    if (filtroActivo.tipo === 'subcategoria') {
      return catalogo.productos.filter((p) => p.subcategoria_id === filtroActivo.subcategoriaId);
    }
    return catalogo.productos;
  }

  function renderGrid() {
    const lista = productosFiltrados();
    gridEl.innerHTML = '';
    statusEl.textContent = `${lista.length} producto${lista.length === 1 ? '' : 's'} encontrado${lista.length === 1 ? '' : 's'}`;

    lista.forEach((producto) => {
      const card = document.createElement('article');
      card.className = 'producto-card';

      const imgWrap = document.createElement('div');
      imgWrap.className = 'producto-img';
      const img = document.createElement('img');
      img.src = imagenPrincipalUrl(producto);
      img.alt = producto.nombre;
      img.loading = 'lazy';
      img.onerror = function () {
        imgWrap.innerHTML = '<div class="img-placeholder">Imagen no disponible</div>';
      };
      imgWrap.appendChild(img);

      const body = document.createElement('div');
      body.className = 'producto-body';
      body.innerHTML = `
        <span class="badge">${producto.subcategoria}</span>
        <div class="producto-titulo">${producto.nombre.toUpperCase()}</div>
        <div class="producto-invima">${producto.registro_invima ? 'Reg. Invima: ' + producto.registro_invima : ''}</div>
        <div class="producto-precio">${formatPrecio(producto.precio)}</div>
      `;

      const btnWa = document.createElement('a');
      btnWa.href = whatsappUrl(producto);
      btnWa.target = '_blank';
      btnWa.rel = 'noopener noreferrer';
      btnWa.className = 'btn-whatsapp';
      btnWa.textContent = 'Comprar por WhatsApp';
      btnWa.addEventListener('click', (e) => e.stopPropagation());
      body.appendChild(btnWa);

      card.appendChild(imgWrap);
      card.appendChild(body);
      card.addEventListener('click', () => abrirModal(producto));

      gridEl.appendChild(card);
    });
  }

  /* ══════════════════════════════════════
     MODAL DE DETALLE + CARRUSEL DINÁMICO
     ══════════════════════════════════════ */
  const modal = document.getElementById('modalProducto');
  const modalImgPrincipal = document.getElementById('modalImgPrincipal');
  const modalThumbs = document.getElementById('modalThumbs');
  const modalBadge = document.getElementById('modalBadge');
  const modalTitulo = document.getElementById('modalTitulo');
  const modalInvima = document.getElementById('modalInvima');
  const modalPrecio = document.getElementById('modalPrecio');
  const modalObs = document.getElementById('modalObs');
  const modalBtnWa = document.getElementById('modalBtnWa');

  async function abrirModal(producto) {
    modalBadge.textContent = producto.subcategoria;
    modalTitulo.textContent = producto.nombre.toUpperCase();
    modalInvima.textContent = producto.registro_invima ? 'Registro Sanitario Invima: ' + producto.registro_invima : '';
    modalPrecio.textContent = formatPrecio(producto.precio);
    modalObs.textContent = producto.observaciones || '';
    modalBtnWa.href = whatsappUrl(producto);

    const principal = imagenPrincipalUrl(producto);
    modalImgPrincipal.src = principal;
    modalImgPrincipal.alt = producto.nombre;
    modalImgPrincipal.onerror = () => { modalImgPrincipal.style.display = 'none'; };
    modalImgPrincipal.onload = () => { modalImgPrincipal.style.display = 'block'; };

    modalThumbs.innerHTML = '';
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';

    // Detecta variantes [Codigo]_01.png, _02.png... y las añade como miniaturas
    const imagenes = [principal];
    for (let n = 1; n <= MAX_VARIANTES; n++) {
      const url = imagenVarianteUrl(producto, n);
      const existe = await checkImagenExiste(url);
      if (!existe) break;
      imagenes.push(url);
    }

    if (imagenes.length > 1) {
      imagenes.forEach((url, idx) => {
        const thumb = document.createElement('img');
        thumb.src = url;
        thumb.alt = producto.nombre + ' miniatura ' + (idx + 1);
        if (idx === 0) thumb.classList.add('active');
        thumb.addEventListener('click', () => {
          modalImgPrincipal.src = url;
          modalThumbs.querySelectorAll('img').forEach((t) => t.classList.remove('active'));
          thumb.classList.add('active');
        });
        modalThumbs.appendChild(thumb);
      });
    }
  }

  function cerrarModal() {
    modal.classList.remove('open');
    document.body.style.overflow = '';
  }

  document.getElementById('modalClose').addEventListener('click', cerrarModal);
  modal.addEventListener('click', (e) => {
    if (e.target === modal) cerrarModal();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modal.classList.contains('open')) cerrarModal();
  });

})();
