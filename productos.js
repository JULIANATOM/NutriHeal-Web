(function () {
  'use strict';

  const BLOB_BASE = 'https://nutriheal.blob.core.windows.net/productos/';
  const WHATSAPP_NUMERO = '573147088080';
  const MAX_VARIANTES = 5;

  let catalogo = { marcas: [], categorias: [], productos: [] };
  let filtroActivo = { marcaId: null, categoriaId: null, subcategoriaId: null };

  const gridEl             = document.getElementById('gridProductos');
  const statusEl           = document.getElementById('gridStatus');
  const listaCategoriasEl  = document.getElementById('listaCategorias');
  const marcaFilterEl      = document.getElementById('marcaFilter');

  /* ── URLs de imagen ──────────────────────────────────────────────────────── */
  // `imagen` en el JSON contiene la ruta relativa al container, incluyendo extensión real.
  function imagenPrincipalUrl(p) {
    if (!p.imagen) return null;
    return BLOB_BASE + p.imagen.split('/').map(encodeURIComponent).join('/');
  }

  function imagenVarianteUrl(p, n) {
    if (!p.imagen) return null;
    const sufijo = String(n).padStart(2, '0');
    // Insertar _01 antes de la extensión: "NV9902.png" → "NV9902_01.png"
    const variante = p.imagen.replace(/(\.[^./]+)$/, `_${sufijo}$1`);
    return BLOB_BASE + variante.split('/').map(encodeURIComponent).join('/');
  }

  function whatsappUrl(p) {
    const texto = `Hola, deseo adquirir el producto ${p.nombre} (Código: ${p.codigo}) de ${p.marca}`;
    return `https://wa.me/${WHATSAPP_NUMERO}?text=${encodeURIComponent(texto)}`;
  }

  function checkImagenExiste(url) {
    return new Promise((resolve) => {
      const img = new Image();
      img.onload  = () => resolve(true);
      img.onerror = () => resolve(false);
      img.src = url;
    });
  }

  /* ── Carga de datos ──────────────────────────────────────────────────────── */
  fetch('catalogo.json')
    .then((r) => {
      if (!r.ok) throw new Error('No se pudo cargar catalogo.json');
      return r.json();
    })
    .then((data) => {
      catalogo = data;
      renderMarcas();
      renderFiltros();
      renderGrid();
    })
    .catch((err) => {
      statusEl.textContent = 'No se pudo cargar el catálogo. Intenta de nuevo más tarde.';
      console.error(err);
    });

  /* ══════════════════════════════════════════════════════════════════════════
     FILTRO DE MARCAS (pills horizontales encima de las categorías)
     ══════════════════════════════════════════════════════════════════════════ */
  function renderMarcas() {
    marcaFilterEl.innerHTML = '';

    const todas = document.createElement('button');
    todas.type = 'button';
    todas.className = 'marca-pill' + (filtroActivo.marcaId === null ? ' active' : '');
    todas.textContent = 'Todas';
    todas.addEventListener('click', () => seleccionarMarca(null));
    marcaFilterEl.appendChild(todas);

    catalogo.marcas.forEach((m) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'marca-pill' + (filtroActivo.marcaId === m.id ? ' active' : '');
      btn.textContent = m.nombre;
      btn.dataset.marcaId = m.id;
      btn.addEventListener('click', () => seleccionarMarca(m.id));
      marcaFilterEl.appendChild(btn);
    });
  }

  function seleccionarMarca(marcaId) {
    filtroActivo = { marcaId, categoriaId: null, subcategoriaId: null };
    renderMarcas();
    renderFiltros();
    renderGrid();
  }

  /* ══════════════════════════════════════════════════════════════════════════
     FILTRO DE CATEGORÍAS / SUBCATEGORÍAS (acordeón)
     ══════════════════════════════════════════════════════════════════════════ */
  function renderFiltros() {
    listaCategoriasEl.innerHTML = '';

    const cats = filtroActivo.marcaId === null
      ? catalogo.categorias
      : catalogo.categorias.filter((c) => c.marca_id === filtroActivo.marcaId);

    cats.forEach((cat) => {
      const item = document.createElement('div');
      item.className = 'cat-item';

      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = 'cat-toggle';
      // Mostrar nombre de marca solo si estamos en vista "Todas"
      const etiqueta = filtroActivo.marcaId === null
        ? `<span>${cat.nombre} <em class="cat-marca">${cat.marca}</em></span><span class="chevron">▶</span>`
        : `<span>${cat.nombre}</span><span class="chevron">▶</span>`;
      toggle.innerHTML = etiqueta;
      toggle.addEventListener('click', () => item.classList.toggle('open'));

      const subList = document.createElement('ul');
      subList.className = 'subcat-list';

      cat.subcategorias.forEach((sub) => {
        const li  = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.textContent = sub.nombre;
        btn.dataset.subcategoriaId = sub.id;
        btn.addEventListener('click', () => {
          filtroActivo = {
            marcaId: filtroActivo.marcaId,
            categoriaId: cat.id,
            subcategoriaId: sub.id,
          };
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
    filtroActivo = { marcaId: filtroActivo.marcaId, categoriaId: null, subcategoriaId: null };
    document.querySelectorAll('.subcat-list button.active').forEach((b) => b.classList.remove('active'));
    renderGrid();
  });

  /* ══════════════════════════════════════════════════════════════════════════
     GRID DE PRODUCTOS
     ══════════════════════════════════════════════════════════════════════════ */
  function productosFiltrados() {
    let lista = catalogo.productos;
    if (filtroActivo.marcaId !== null) {
      lista = lista.filter((p) => p.marca_id === filtroActivo.marcaId);
    }
    if (filtroActivo.subcategoriaId !== null) {
      lista = lista.filter((p) => p.subcategoria_id === filtroActivo.subcategoriaId);
    }
    return lista;
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
      const url = imagenPrincipalUrl(producto);
      if (url) {
        const img = document.createElement('img');
        img.src = url;
        img.alt = producto.nombre;
        img.loading = 'lazy';
        img.onerror = function () {
          imgWrap.innerHTML = '<div class="img-placeholder">Imagen no disponible</div>';
        };
        imgWrap.appendChild(img);
      } else {
        imgWrap.innerHTML = '<div class="img-placeholder">Imagen no disponible</div>';
      }

      const body = document.createElement('div');
      body.className = 'producto-body';
      body.innerHTML = `
        <span class="badge">${producto.subcategoria}</span>
        <div class="producto-titulo">${producto.nombre.toUpperCase()}</div>
        <div class="producto-marca-tag">${producto.marca}</div>
        <div class="producto-invima">${producto.registro_invima ? 'Reg. Invima: ' + producto.registro_invima : ''}</div>
      `;

      const btnWa = document.createElement('a');
      btnWa.href   = whatsappUrl(producto);
      btnWa.target = '_blank';
      btnWa.rel    = 'noopener noreferrer';
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

  /* ══════════════════════════════════════════════════════════════════════════
     MODAL DE DETALLE + CARRUSEL DINÁMICO
     ══════════════════════════════════════════════════════════════════════════ */
  const modal             = document.getElementById('modalProducto');
  const modalImgPrincipal = document.getElementById('modalImgPrincipal');
  const modalThumbs       = document.getElementById('modalThumbs');
  const modalBadge        = document.getElementById('modalBadge');
  const modalTitulo       = document.getElementById('modalTitulo');
  const modalInvima       = document.getElementById('modalInvima');
  const modalPrecio       = document.getElementById('modalPrecio');
  const modalObs          = document.getElementById('modalObs');
  const modalBtnWa        = document.getElementById('modalBtnWa');

  async function abrirModal(producto) {
    modalBadge.textContent  = producto.subcategoria;
    modalTitulo.textContent = producto.nombre.toUpperCase();
    modalInvima.textContent = producto.registro_invima
      ? 'Registro Sanitario Invima: ' + producto.registro_invima
      : '';
    modalPrecio.textContent = '';
    modalObs.textContent    = producto.observaciones || '';
    modalBtnWa.href         = whatsappUrl(producto);

    const principal = imagenPrincipalUrl(producto);
    if (principal) {
      modalImgPrincipal.src     = principal;
      modalImgPrincipal.style.display = 'block';
    } else {
      modalImgPrincipal.style.display = 'none';
    }
    modalImgPrincipal.alt     = producto.nombre;
    modalImgPrincipal.onerror = () => { modalImgPrincipal.style.display = 'none'; };
    modalImgPrincipal.onload  = () => { modalImgPrincipal.style.display = 'block'; };

    modalThumbs.innerHTML = '';
    modal.classList.add('open');
    document.body.style.overflow = 'hidden';

    const imagenes = principal ? [principal] : [];
    for (let n = 1; n <= MAX_VARIANTES; n++) {
      const url = imagenVarianteUrl(producto, n);
      if (!url) break;
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
  modal.addEventListener('click', (e) => { if (e.target === modal) cerrarModal(); });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && modal.classList.contains('open')) cerrarModal();
  });

})();
