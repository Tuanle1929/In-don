/**
 * =========================================================================
 * BACKEND GOOGLE APPS SCRIPT - TẠO ĐƠN ĐẶT HÀNG TOBOKO (v1.3)
 * Cột tab SanPham: MaSP | TenSP | PhanLoai | DVT | DonGia
 * =========================================================================
 */

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const schemas = {
    'SanPham': ['MaSP', 'TenSP', 'PhanLoai', 'DVT', 'DonGia'],
    'KhachHang': ['SDT', 'TenKhach', 'DiaChi', 'CapNhatCuoi'],
    'DonHang': ['MaDon', 'NgayLap', 'TenKhach', 'SDT', 'DiaChi', 'ChiTietMon_JSON', 'TongTienHang', 'PhiShip', 'TongCong', 'GhiChu']
  };

  for (const [sheetName, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center');
      sheet.setFrozenRows(1);
    } else if (sheetName === 'SanPham') {
      const lastCol = Math.max(sheet.getLastColumn(), 1);
      const firstRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (!firstRow.includes('PhanLoai')) {
        sheet.insertColumnAfter(2);
        sheet.getRange(1, 3)
          .setValue('PhanLoai')
          .setFontWeight('bold')
          .setBackground('#1e293b')
          .setFontColor('#ffffff')
          .setHorizontalAlignment('center');
      }
    }
  }
}

function parsePriceNumber(val) {
  if (typeof val === 'number') return Math.round(val);
  if (!val) return 0;
  const cleaned = String(val).replace(/[^\d]/g, '');
  return parseInt(cleaned, 10) || 0;
}

/**
 * API GET: Tải danh mục Sản Phẩm, Khách Hàng, Đơn Hàng
 */
function doGet(e) {
  setupSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const getRows = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return [];
    const headers = rows[0];
    return rows.slice(1).map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });
  };

  // 1. Đọc danh mục sản phẩm & hỗ trợ tách dấu phẩy trong cột PhanLoai
  const rawProducts = getRows('SanPham');
  const products = [];

  rawProducts.forEach(p => {
    const name = String(p.TenSP || '').trim();
    if (!name) return;
    const code = String(p.MaSP || '').trim();
    const rawVariant = String(p.PhanLoai || '').trim();
    const unit = String(p.DVT || 'Hộp').trim();
    const price = parsePriceNumber(p.DonGia);

    if (rawVariant.includes(',')) {
      const variants = rawVariant.split(',').map(v => v.trim()).filter(Boolean);
      variants.forEach(v => {
        products.push({
          code: code ? (code + '-' + v) : v,
          name: name,
          variant: v,
          unit: unit,
          price: price
        });
      });
    } else {
      products.push({
        code: code,
        name: name,
        variant: rawVariant,
        unit: unit,
        price: price
      });
    }
  });

  // 2. Danh bạ khách hàng
  const customers = getRows('KhachHang').map(c => ({
    phone: String(c.SDT || '').trim(),
    name: String(c.TenKhach || '').trim(),
    address: String(c.DiaChi || '').trim(),
    updatedAt: String(c.CapNhatCuoi || '')
  })).filter(c => c.name || c.phone);

  // 3. Đơn hàng gần nhất
  const rawOrders = getRows('DonHang').slice(-100).reverse();
  const orders = rawOrders.map(o => {
    let items = [];
    try { 
      items = JSON.parse(o.ChiTietMon_JSON || '[]'); 
    } catch (_) {
      items = [];
    }
    const rawShip = String(o.PhiShip || '').trim();
    let shipType = 'amount';
    if (rawShip === 'free' || rawShip.toLowerCase().includes('miễn phí')) {
      shipType = 'free';
    } else if (rawShip === 'cod_ship' || rawShip.toLowerCase().includes('tự trả')) {
      shipType = 'cod_ship';
    }
    return {
      id: String(o.MaDon || ''),
      createdAt: String(o.NgayLap || ''),
      customer: {
        name: String(o.TenKhach || ''),
        phone: String(o.SDT || ''),
        address: String(o.DiaChi || '')
      },
      items: items,
      subtotal: parsePriceNumber(o.TongTienHang),
      shipFee: parsePriceNumber(o.PhiShip),
      shipType: shipType,
      total: parsePriceNumber(o.TongCong),
      note: String(o.GhiChu || '')
    };
  });

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, products, customers, orders }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * API POST: Lưu/Cập nhật Đơn Hàng + Upsert Khách Hàng + Đồng bộ Sản Phẩm
 */
function doPost(e) {
  setupSheets();
  try {
    const contents = e.postData ? e.postData.contents : '{}';
    const payload = JSON.parse(contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.action === 'SAVE_ORDER' && payload.order) {
      const o = payload.order;
      const orderSheet = ss.getSheetByName('DonHang');
      const data = orderSheet.getDataRange().getValues();
      
      const phoneFormatted = o.customer.phone ? ("'" + o.customer.phone) : '';
      const shipCol = (o.shipType === 'amount') ? Number(o.shipFee || 0) : (o.shipType === 'free' ? 'Miễn phí ship' : 'Khách tự trả ship');

      const rowData = [
        String(o.id),
        String(o.createdAt),
        String(o.customer.name || ''),
        phoneFormatted,
        String(o.customer.address || ''),
        JSON.stringify(o.items || []),
        Number(o.subtotal || 0),
        shipCol,
        Number(o.total || 0),
        String(o.note || '')
      ];

      let updated = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === String(o.id).trim()) {
          orderSheet.getRange(i + 1, 1, 1, rowData.length).setValues([rowData]);
          updated = true;
          break;
        }
      }
      if (!updated) {
        orderSheet.appendRow(rowData);
      }

      if (o.customer && (o.customer.name || o.customer.phone)) {
        upsertCustomer(ss, o.customer);
      }
    }

    if (payload.action === 'SYNC_PRODUCTS' && Array.isArray(payload.products)) {
      const prodSheet = ss.getSheetByName('SanPham');
      prodSheet.clearContents();
      prodSheet.appendRow(['MaSP', 'TenSP', 'PhanLoai', 'DVT', 'DonGia']);
      prodSheet.getRange(1, 1, 1, 5)
        .setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center');
      prodSheet.setFrozenRows(1);

      if (payload.products.length > 0) {
        const rows = payload.products.map(p => [
          String(p.code || ''),
          String(p.name || ''),
          String(p.variant || p.category || ''),
          String(p.unit || 'Hộp'),
          parsePriceNumber(p.price)
        ]);
        prodSheet.getRange(2, 1, rows.length, 5).setValues(rows);
      }
    }

    if (payload.action === 'SYNC_CUSTOMERS' && Array.isArray(payload.customers)) {
      const custSheet = ss.getSheetByName('KhachHang');
      custSheet.clearContents();
      custSheet.appendRow(['SDT', 'TenKhach', 'DiaChi', 'CapNhatCuoi']);
      custSheet.getRange(1, 1, 1, 4)
        .setFontWeight('bold')
        .setBackground('#1e293b')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center');
      custSheet.setFrozenRows(1);

      if (payload.customers.length > 0) {
        const nowStr = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
        const rows = payload.customers.map(c => [
          c.phone ? ("'" + c.phone) : '',
          String(c.name || ''),
          String(c.address || ''),
          c.updatedAt || nowStr
        ]);
        custSheet.getRange(2, 1, rows.length, 4).setValues(rows);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, message: 'Đồng bộ thành công!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function upsertCustomer(ss, customer) {
  const custSheet = ss.getSheetByName('KhachHang');
  const cData = custSheet.getDataRange().getValues();
  const cleanPhone = String(customer.phone || '').replace(/\D/g, '');
  const custName = String(customer.name || '').trim().toLowerCase();
  const phoneFormatted = customer.phone ? ("'" + customer.phone) : '';
  const nowStr = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');

  for (let i = 1; i < cData.length; i++) {
    const rowPhone = String(cData[i][0] || '').replace(/\D/g, '');
    const rowName = String(cData[i][1] || '').trim().toLowerCase();
    
    if ((cleanPhone && rowPhone === cleanPhone) || (!cleanPhone && rowName === custName)) {
      custSheet.getRange(i + 1, 1, 1, 4).setValues([[
        phoneFormatted,
        customer.name,
        customer.address,
        nowStr
      ]]);
      return;
    }
  }

  custSheet.appendRow([
    phoneFormatted,
    customer.name,
    customer.address,
    nowStr
  ]);
}
