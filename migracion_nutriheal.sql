-- ============================================================
-- MIGRACION NUTRIHEAL 360 -- Esquema + Carga de Datos
-- Base de datos: ERP_DB_PROD | Esquema: nutriheal
-- Generado automaticamente desde Bd.xlsx (Hoja1)
-- ============================================================

-- 1. ESQUEMA
IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = 'nutriheal')
    EXEC('CREATE SCHEMA nutriheal');
GO

-- 2. TABLAS

IF OBJECT_ID('nutriheal.productos', 'U') IS NOT NULL DROP TABLE nutriheal.productos;
IF OBJECT_ID('nutriheal.subcategorias', 'U') IS NOT NULL DROP TABLE nutriheal.subcategorias;
IF OBJECT_ID('nutriheal.categorias', 'U') IS NOT NULL DROP TABLE nutriheal.categorias;
GO

CREATE TABLE nutriheal.categorias (
    id      INT IDENTITY(1,1) PRIMARY KEY,
    nombre  NVARCHAR(100) NOT NULL UNIQUE
);
GO

CREATE TABLE nutriheal.subcategorias (
    id            INT IDENTITY(1,1) PRIMARY KEY,
    categoria_id  INT NOT NULL REFERENCES nutriheal.categorias(id),
    nombre        NVARCHAR(100) NOT NULL,
    CONSTRAINT UQ_subcategoria UNIQUE (categoria_id, nombre)
);
GO

CREATE TABLE nutriheal.productos (
    id                INT IDENTITY(1,1) PRIMARY KEY,
    codigo            NVARCHAR(20) NOT NULL UNIQUE,
    nombre            NVARCHAR(200) NOT NULL,
    subcategoria_id   INT NOT NULL REFERENCES nutriheal.subcategorias(id),
    registro_invima   NVARCHAR(50) NULL,
    precio            DECIMAL(10,2) NOT NULL DEFAULT 0,
    observaciones     NVARCHAR(MAX) NULL,
    activo            BIT NOT NULL DEFAULT 1
);
GO

-- 3. CARGA: CATEGORIAS
INSERT INTO nutriheal.categorias (nombre) VALUES
    (N'Suplementos dietarios'),
    (N'Medicamentos'),
    (N'Fitoterapéuticos');
GO

-- 4. CARGA: SUBCATEGORIAS
INSERT INTO nutriheal.subcategorias (categoria_id, nombre) VALUES
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Gastrointestinal'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Probiotics'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'D-Tox'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Hepático'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Cardiovascular'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Insuficiencia Venosa'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Articular'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Medicamentos'), N'Articular'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Fitoterapéuticos'), N'Articular'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Sistema Inmunológico'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Metabolismo'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Fitoterapéuticos'), N'Sistema Nervioso'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Sistema Nervioso'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Medicamentos'), N'Sistema Nervioso'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Urinario'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Medicamentos'), N'P-Natal'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Visión'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Endocrino'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Antioxidantes'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Dermatología'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Figura'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Minerales'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Vitaminas'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Multivitamínicos'),
    ((SELECT id FROM nutriheal.categorias WHERE nombre = N'Suplementos dietarios'), N'Masa Muscular');
GO

-- 5. CARGA: PRODUCTOS
INSERT INTO nutriheal.productos (codigo, nombre, subcategoria_id, registro_invima, precio, observaciones) VALUES
    (N'NV9902', N'BETAPEPZINC', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Gastrointestinal'), N'SD2021-0004596', 0, NULL),
    (N'NV1713', N'DIGEZIMES', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Gastrointestinal'), N'SD2015-0003438', 0, NULL),
    (N'NV1073', N'SPOROGENEX', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Probiotics'), N'SD2023-0004692', 0, NULL),
    (N'NV1767', N'FLORAWELL', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Probiotics'), N'SD2013-0002822', 0, NULL),
    (N'NV468', N'VOID', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'D-Tox'), N'SD2012-0002604', 0, NULL),
    (N'NV467', N'D-PURE AM-PM', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'D-Tox'), N'SD2013-0003014', 0, NULL),
    (N'NV1689', N'LAX-ON', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'D-Tox'), N'SD2013-0002888', 0, NULL),
    (N'NV1682', N'LEBER', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Hepático'), N'SD2013-0002935', 0, NULL),
    (N'NV1975', N'HEPINÉ', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Hepático'), N'PFTI2021-0002815', 0, NULL),
    (N'NV1671', N'MEGA RED KRILL', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Cardiovascular'), N'SD2013-0002984', 0, NULL),
    (N'NV458', N'OCEAN BLUE', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Cardiovascular'), N'SD2021-0004575', 0, NULL),
    (N'NV1692', N'CITRICHOLESS', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Cardiovascular'), N'SD2013-0002907', 0, NULL),
    (N'NV2124', N'C-ZYME Q10', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Cardiovascular'), N'SD2020-0004508', 0, NULL),
    (N'NV1548', N'CORONALL', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Cardiovascular'), N'SD2024-0004734', 0, NULL),
    (N'NV1728', N'NEO VEINS', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Insuficiencia Venosa'), N'INVIMA 2015M-0016530', 0, N'* Venta con fórmula facultativa. * No exceda su consumo. * Leer indicaciones y contraindicaciones en la etiqueta. * Si los síntomas persisten, consultar al médico.'),
    (N'NV9949', N'OSSEIM UCII UNDERNATURE COLLAGEN TYPE II', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Articular'), N'SD2015-0003597', 0, NULL),
    (N'NV1546A', N'NUTRA JOINT COMFORT', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Medicamentos' AND s.nombre = N'Articular'), N'INVIMA 2018M-0018048', 0, N'* Venta con fórmula facultativa. * No exceda su consumo. * Leer indicaciones y contraindicaciones en la etiqueta. * Si los síntomas persisten, consultar al médico.'),
    (N'NV1691', N'DYPHLAMIN', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Fitoterapéuticos' AND s.nombre = N'Articular'), N'PFM2012-0001934', 0, N'* De venta libre. * No exceda su consumo. * Leer indicaciones, contraindicaciones y advertencias en la etiqueta. * Si los síntomas persisten, consultar al médico.'),
    (N'NV1683', N'IMMBOOST', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Sistema Inmunológico'), N'SD2013-0002758', 0, NULL),
    (N'NV2074', N'ISOTYPES 5 IMG 1.000 MG', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Sistema Inmunológico'), N'SD2016-0003776', 0, NULL),
    (N'NV1686', N'ALLER-7', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Sistema Inmunológico'), N'SD2015-0003460', 0, NULL),
    (N'NV1788', N'GLUCOVITA', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Metabolismo'), N'SD2013-0002970', 0, NULL),
    (N'NV1706', N'CHROMOX GTF CHROMIUM PICOLINATE 500mg', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Metabolismo'), N'SD2012-0002630', 0, NULL),
    (N'NV1976', N'DYROX', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Metabolismo'), N'SD2015-0003690', 0, NULL),
    (N'NV1684', N'DEMZENIL', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Fitoterapéuticos' AND s.nombre = N'Sistema Nervioso'), N'PFM2012-0001915', 0, N'* Venta con fórmula facultativa. * No exceda su consumo. * Leer indicaciones y contraindicaciones en la etiqueta. * Si los síntomas persisten, consultar al médico.'),
    (N'NV1693', N'MEMOBRITE', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Sistema Nervioso'), N'SD2013-0002832', 0, NULL),
    (N'NV759', N'SOMNA', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Sistema Nervioso'), N'SD2013-0002889', 0, NULL),
    (N'NV835', N'NEUROX', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Sistema Nervioso'), N'SD2014-0001864-R1', 0, NULL),
    (N'NV1670', N'MELATONINA', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Medicamentos' AND s.nombre = N'Sistema Nervioso'), N'INVIMA 2015M-0016375', 0, N'* Venta con fórmula facultativa. * No exceda su consumo. * Leer indicaciones y contraindicaciones en la etiqueta. * Si los síntomas persisten, consultar al médico.'),
    (N'NV1690', N'PROGENIL', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Urinario'), N'SD2013-0002847', 0, NULL),
    (N'NV1911', N'MAXCRAN CRANBERRY 1.100 MG', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Urinario'), N'SD2015-0003680', 0, NULL),
    (N'NV1583', N'NEFROX', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Urinario'), N'SD2013-0002992', 0, NULL),
    (N'NV1904', N'P-NATAL WITH DHA', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Medicamentos' AND s.nombre = N'P-Natal'), N'INVIMA 2016M-0016999', 0, NULL),
    (N'NV2046', N'VITAVER', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Visión'), N'SD2015-0003712', 0, NULL),
    (N'NV1685', N'GENIX', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Endocrino'), N'SD2013-0002901', 0, NULL),
    (N'NV1981', N'SURFEM', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Endocrino'), N'SD2014-0003324', 0, NULL),
    (N'NV1688', N'FemmeSense', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Endocrino'), N'SD2013-0002864', 0, NULL),
    (N'NV1886', N'TORNARE', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Antioxidantes'), N'SD2015-0003674', 0, NULL),
    (N'NV9901', N'LP-SOMAL L-GLUTATHIONE 500MG', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Antioxidantes'), N'SD2021-0004601', 0, NULL),
    (N'NV1982', N'KAPILSAN', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Dermatología'), N'SD2015-0003659', 0, NULL),
    (N'NV1992', N'XENSOR', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Figura'), N'SD2018-0004167', 0, NULL),
    (N'NV1667', N'OSTEOWELL', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Minerales'), N'SD2011-0002155', 0, NULL),
    (N'NV1795', N'OSEOMASS CAL-MAG-ZINC PLUS VD3', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Minerales'), N'SD2014-0003260', 0, NULL),
    (N'NV1875', N'NOPHOS MAGNESIUM CITRATE', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Minerales'), N'SD2025-0004968', 0, NULL),
    (N'NV1876', N'MAGX3', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Minerales'), N'SD2015-0003666', 0, NULL),
    (N'NV197', N'CAROTENALL E', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Vitaminas'), N'SD2025-0004956', 0, NULL),
    (N'NV344', N'METHYLCOBALAMIN B-12', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Vitaminas'), N'SD2016-0003746', 0, NULL),
    (N'NV290', N'B-100 ADVANCED', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Vitaminas'), N'SD2021-0001561-R1', 0, NULL),
    (N'NV1942', N'VITAMINA C', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Vitaminas'), N'SD2015-0003704', 0, NULL),
    (N'NV1874', N'FEROL D3 2.000 UI', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Vitaminas'), N'SD2015-0003523', 0, NULL),
    (N'NV2134', N'OSTIK2 (VITAMIN D3+K2)', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Vitaminas'), N'SD2021-0004638', 0, NULL),
    (N'NV1980', N'RE-ACTIVIN', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Multivitamínicos'), N'SD2016-0003778', 0, NULL),
    (N'NV1680A', N'GERIAFULL', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Multivitamínicos'), N'SD2012-0002402', 0, NULL),
    (N'NV1774', N'TOTAL DAY', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Multivitamínicos'), N'SD2013-0003012', 0, NULL),
    (N'NV1678', N'KIDDY MULTYVITAMIN', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Multivitamínicos'), N'SD2013-0002962', 0, NULL),
    (N'NV1674', N'WOMEN´S COMPLETE', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Multivitamínicos'), N'SD2012-0002667', 0, NULL),
    (N'NV1700', N'MEN´S MEGA MULTI', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Multivitamínicos'), N'SD2013-0002695', 0, NULL),
    (N'NV9976', N'VEGAN PROTEIN', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Masa Muscular'), N'RSA-001298-2016', 0, NULL),
    (N'NV9929', N'BUTY (HMB COMPLEX)', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Masa Muscular'), N'SD2024-0004771', 0, NULL),
    (N'NV9937', N'DISFUTIL- L-ARGININE 1,000MG', (SELECT s.id FROM nutriheal.subcategorias s JOIN nutriheal.categorias c ON c.id = s.categoria_id WHERE c.nombre = N'Suplementos dietarios' AND s.nombre = N'Masa Muscular'), N'SD2017-0004098', 0, NULL);
GO

-- 6. VERIFICACION
SELECT 'categorias' AS tabla, COUNT(*) AS registros FROM nutriheal.categorias
UNION ALL SELECT 'subcategorias', COUNT(*) FROM nutriheal.subcategorias
UNION ALL SELECT 'productos', COUNT(*) FROM nutriheal.productos;
GO