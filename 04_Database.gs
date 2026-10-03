const MC_Database = {
  ensureSheet: function(n, h) {
    var ss = this.ss();
    var s = ss.getSheetByName(n);
    if (!s) {
      s = ss.insertSheet(n);
      s.getRange(1, 1, 1, h.length).setValues([h]);
      s.setFrozenRows(1);
    } else if (s.getLastRow() === 0) {
      s.getRange(1, 1, 1, h.length).setValues([h]);
      s.setFrozenRows(1);
    }
    return s;
  },

  ss: function() {
    var id = PropertiesService.getScriptProperties()
      .getProperty(MC_CONFIG.DB_PROPERTY);

    if (id) {
      try {
        return SpreadsheetApp.openById(id);
      } catch (e) {
        // Fall through and recover from the bound spreadsheet.
      }
    }

    var active = SpreadsheetApp.getActiveSpreadsheet();
    MC_Utils.require(active, 'Database belum di-setup');
    PropertiesService.getScriptProperties()
      .setProperty(MC_CONFIG.DB_PROPERTY, active.getId());

    return active;
  },

  sheet: function(n) {
    var s = this.ss().getSheetByName(n);
    MC_Utils.require(s, 'Sheet ' + n + ' tidak ditemukan');
    return s;
  },

  rows: function(n) {
    var v = this.sheet(n).getDataRange().getValues();
    if (v.length < 2) return [];
    return v.slice(1)
      .filter(function(r) {
        return r.some(function(x) {
          return x !== '' && x !== null;
        });
      })
      .map(function(r) {
        return Object.fromEntries(
          v[0].map(function(k, i) {
            return [k, r[i]];
          })
        );
      });
  },

  find: function(n, k, v) {
    return this.rows(n).find(function(x) {
      return String(x[k]) === String(v);
    }) || null;
  },

  update: function(n, k, v, o) {
    var s = this.sheet(n);
    var data = s.getDataRange().getValues();
    var h = data[0];
    var idx = h.indexOf(k);

    MC_Utils.require(idx >= 0, 'Kolom ' + k + ' tidak ditemukan');

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][idx]) === String(v)) {
        s.getRange(r + 1, 1, 1, h.length).setValues([
          h.map(function(key, i) {
            return o[key] ?? data[r][i];
          })
        ]);
        return o;
      }
    }

    throw new Error('Record tidak ditemukan');
  },

  insert: function(n, o) {
    var s = this.sheet(n);
    var h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
    s.appendRow(h.map(function(k) {
      return o[k] ?? '';
    }));
    return o;
  },

  removeWhere: function(n, k, v) {
    var s = this.sheet(n);
    var data = s.getDataRange().getValues();
    var h = data[0];
    var idx = h.indexOf(k);

    MC_Utils.require(idx >= 0, 'Kolom ' + k + ' tidak ditemukan');

    for (var r = data.length - 1; r >= 1; r--) {
      if (String(data[r][idx]) === String(v)) {
        s.deleteRow(r + 1);
      }
    }
    return true;
  },

  audit: function(m, a, id, o, n, r) {
    this.insert('AUDIT_LOG', {
      ID: Utilities.getUuid(),
      USER_ID: Session.getActiveUser().getEmail() || 'system',
      TIMESTAMP: new Date(),
      MODULE: m,
      ACTION: a,
      RECORD_ID: id,
      OLD_VALUE: JSON.stringify(o ?? null),
      NEW_VALUE: JSON.stringify(n ?? null),
      IP: '',
      SESSION_ID: Session.getTemporaryActiveUserKey(),
      REASON: r || ''
    });
  },

  migrateSchema_: function(ss) {
    var changed=[];
    Object.keys(MC_SHEETS).forEach(function(n){
      var s=ss.getSheetByName(n);
      if(!s)return;
      var desired=MC_SHEETS[n];
      var lastCol=Math.max(s.getLastColumn(),1);
      var current=s.getRange(1,1,1,lastCol).getValues()[0].map(function(x){return String(x||'').trim()});
      if(!current.some(function(x){return x!=='';})){
        s.getRange(1,1,1,desired.length).setValues([desired]);
        changed.push({sheet:n,added:desired});
        return;
      }
      var missing=desired.filter(function(h){return current.indexOf(h)<0});
      if(missing.length){
        var start=current.length+1;
        s.getRange(1,start,1,missing.length).setValues([missing]);
        changed.push({sheet:n,added:missing});
      }
    });
    return changed;
  },

  install: function() {
    var start = Date.now();
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty(MC_CONFIG.DB_PROPERTY);
    var ss = null;

    if (id) {
      try {
        ss = SpreadsheetApp.openById(id);
      } catch (e) {
        ss = null;
      }
    }

    // This project is spreadsheet-bound. Prefer the bound spreadsheet so
    // installation never silently creates a second database.
    if (!ss) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }

    MC_Utils.require(ss, 'Spreadsheet database tidak ditemukan');
    props.setProperty(MC_CONFIG.DB_PROPERTY, ss.getId());

    var created = 0;
    var initialized = 0;

    // Create/initialize all sheets with one header write per sheet.
    Object.keys(MC_SHEETS).forEach(function(n) {
      var s = ss.getSheetByName(n);

      if (!s) {
        s = ss.insertSheet(n);
        created++;
      }

      var h = MC_SHEETS[n];

      if (s.getLastRow() === 0) {
        s.getRange(1, 1, 1, h.length).setValues([h]);
        initialized++;
      } else {
        var current = s.getRange(1, 1, 1, Math.max(s.getLastColumn(), 1))
          .getValues()[0];

        // Empty/placeholder first row: initialize it.
        if (!current.some(function(x) {
          return x !== '' && x !== null;
        })) {
          s.getRange(1, 1, 1, h.length).setValues([h]);
          initialized++;
        }
      }

      s.setFrozenRows(1);
    });

    SpreadsheetApp.flush();

    var migrations = this.migrateSchema_(ss);
    var schemaSheet = ss.getSheetByName('CONFIG');
    var schemaExists = schemaSheet.getDataRange().getValues().slice(1).some(function(r){return String(r[0]||'')==='SCHEMA_VERSION'});
    if(!schemaExists){schemaSheet.getRange(schemaSheet.getLastRow()+1,1,1,3).setValues([['SCHEMA_VERSION','8.1','MC-Almara schema migration version']]);}
    SpreadsheetApp.flush();

    var seedResult = this.seedBatch_(ss);

    SpreadsheetApp.flush();

    var result = {
      ok: true,
      spreadsheetId: ss.getId(),
      url: ss.getUrl(),
      sheetCount: Object.keys(MC_SHEETS).length,
      createdSheets: created,
      initializedSheets: initialized,
      rolesInserted: seedResult.rolesInserted,
      permissionsInserted: seedResult.permissionsInserted,
      configInserted: seedResult.configInserted,
      currenciesInserted: seedResult.currenciesInserted,
      migrations: migrations,
      elapsedMs: Date.now() - start
    };

    console.log('MC-ALMARA INSTALL COMPLETE', result);
    return result;
  },

  seedBatch_: function(ss) {
    var result = {
      rolesInserted: 0,
      permissionsInserted: 0,
      configInserted: 0,
      currenciesInserted: 0
    };

    // Roles: read once, append missing rows in one batch.
    var roleSheet = ss.getSheetByName('ROLES');
    var roleData = roleSheet.getDataRange().getValues();
    var roleHeader = roleData[0];
    var existingRoles = roleData.slice(1).map(function(r) {
      return String(r[0] || '');
    });

    var roleRows = [];
    MC_CONFIG.ROLES.forEach(function(name, i) {
      var id = 'ROLE-' + (i + 1);
      if (existingRoles.indexOf(id) < 0) {
        roleRows.push([
          id,
          name,
          '',
          'ACTIVE'
        ]);
      }
    });

    if (roleRows.length) {
      roleSheet.getRange(
        roleSheet.getLastRow() + 1,
        1,
        roleRows.length,
        roleHeader.length
      ).setValues(roleRows);
      result.rolesInserted = roleRows.length;
    }

    // Permissions: build the complete desired matrix in memory,
    // compare once, then append all missing rows in one write.
    var permSheet = ss.getSheetByName('PERMISSIONS');
    var permData = permSheet.getDataRange().getValues();
    var permHeader = permData[0];
    var existingPerms = {};

    permData.slice(1).forEach(function(r) {
      var key = [
        String(r[1] || ''),
        String(r[2] || ''),
        String(r[3] || '')
      ].join('|');
      existingPerms[key] = true;
    });

    var matrix = {
      Admin: [
        'view','create','edit','delete','print','export','approve',
        'closing','rate_management','transaction_edit','wa'
      ],
      Manager: [
        'view','create','edit','print','export','approve',
        'closing','rate_management','transaction_edit','wa'
      ],
      Teller: [
        'view','create','edit','print','transaction_edit','wa'
      ],
      Kasir: [
        'view','create','edit','print','wa'
      ],
      Finance: [
        'view','create','edit','print','export','approve','closing'
      ],
      Auditor: [
        'view','print','export'
      ]
    };

    var permRows = [];

    MC_CONFIG.ROLES.forEach(function(name, i) {
      var roleId = 'ROLE-' + (i + 1);
      var permissions = matrix[name] || ['view'];

      MC_CONFIG.MODULES.forEach(function(module) {
        permissions.forEach(function(permission) {
          var key = [
            roleId,
            permission,
            module
          ].join('|');

          if (!existingPerms[key]) {
            permRows.push([
              Utilities.getUuid(),
              roleId,
              permission,
              module,
              'ACTIVE'
            ]);
            existingPerms[key] = true;
          }
        });
      });
    });

    if (permRows.length) {
      permSheet.getRange(
        permSheet.getLastRow() + 1,
        1,
        permRows.length,
        permHeader.length
      ).setValues(permRows);
      result.permissionsInserted = permRows.length;
    }

    // Config: seed only when the key is absent.
    var configSheet = ss.getSheetByName('CONFIG');
    var configData = configSheet.getDataRange().getValues();
    var hasThreshold = configData.slice(1).some(function(r) {
      return String(r[0] || '') === 'TRANSACTION_THRESHOLD_USD';
    });

    if (!hasThreshold) {
      configSheet.getRange(configSheet.getLastRow() + 1, 1, 1, 3)
        .setValues([[
          'TRANSACTION_THRESHOLD_USD',
          '10000',
          'Monthly customer transaction threshold in USD equivalent'
        ]]);
      result.configInserted++;
    }

    // Currency master: seed only when empty.
    var currencySheet = ss.getSheetByName('CURRENCY_MASTER');
    var currencyData = currencySheet.getDataRange().getValues();
    var hasCurrencyRows = currencyData.length > 1 &&
      currencyData.slice(1).some(function(r) {
        return r.some(function(x) {
          return x !== '' && x !== null;
        });
      });

    if (!hasCurrencyRows) {
      var currencies = [
        ['CUR-1','USD','US Dollar','', '', '', 2, 'ACTIVE', 1],
        ['CUR-2','EUR','Euro','', '', '', 2, 'ACTIVE', 2],
        ['CUR-3','JPY','Japanese Yen','', '', '', 0, 'ACTIVE', 3],
        ['CUR-4','GBP','Pound Sterling','', '', '', 2, 'ACTIVE', 4],
        ['CUR-5','SGD','Singapore Dollar','', '', '', 2, 'ACTIVE', 5],
        ['CUR-6','AUD','Australian Dollar','', '', '', 2, 'ACTIVE', 6],
        ['CUR-7','CNY','Chinese Yuan','', '', '', 2, 'ACTIVE', 7]
      ];

      currencySheet.getRange(
        currencySheet.getLastRow() + 1,
        1,
        currencies.length,
        currencies[0].length
      ).setValues(currencies);

      result.currenciesInserted = currencies.length;
    }

    return result;
  },

  seed: function() {
    // Backward-compatible entry point.
    return this.seedBatch_(this.ss());
  }
};
