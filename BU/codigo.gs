// =========================================================================
// CONFIGURACIÓ GLOBAL I SEGURETAT D'EDITORS
// =========================================================================
var SPREADSHEET_ID = "1T7K0skPnyhE8ZivKbk1QLduhJwhDrQvkUOZIpvas4yI";

var EDITORS_PRICING = [
  "miquel.navarro@grupserhs.com",
  "comercial.hotels@grupserhs.com"
];

function verificarPermisEdicioPricing() {
  var emailActiu = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!emailActiu) return false;
  return EDITORS_PRICING.indexOf(emailActiu) !== -1;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Panell Revenue & Pricing")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// =========================================================================
// UTILITATS DE FORMAT DE DATA
// =========================================================================
function formatarDataISO(data) {
  if (!data) return "";
  if (data instanceof Date && !isNaN(data.getTime())) {
    var a = data.getFullYear();
    var m = data.getMonth() + 1;
    var d = data.getDate();
    return a + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);
  }
  var str = String(data).trim();
  if (!str) return "";
  if (str.indexOf("T") !== -1) str = str.split("T")[0];
  var parts = str.split(/[\/\-\.]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      var a = parts[0];
      var m = parts[1].length === 1 ? "0" + parts[1] : parts[1];
      var d = parts[2].length === 1 ? "0" + parts[2] : parts[2];
      return a + "-" + m + "-" + d;
    } else {
      var d = parts[0].length === 1 ? "0" + parts[0] : parts[0];
      var m = parts[1].length === 1 ? "0" + parts[1] : parts[1];
      var a = parts[2];
      if (a.length === 2) a = "20" + a;
      return a + "-" + m + "-" + d;
    }
  }
  return "";
}

function restarDies(dataStr, n) {
  if (!dataStr) return "";
  var parts = dataStr.split("-");
  var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  d.setDate(d.getDate() - n);
  return formatarDataISO(d);
}

function obtenirDataNaturalAnyAnterior(dataStr) {
  if (!dataStr || dataStr.length < 10) return "";
  var parts = dataStr.split("-");
  var anyLY = parseInt(parts[0], 10) - 1;
  return anyLY + "-" + parts[1] + "-" + parts[2];
}

function obtenirCapacitatHotel(nom) {
  if (!nom) return 100;
  var n = String(nom).trim().toLowerCase();
  if (n.indexOf("rivoli") !== -1) return 126;
  if (n.indexOf("port") !== -1) return 48;
  if (n.indexOf("carlit") !== -1) return 38;
  if (n.indexOf("hilari") !== -1) return 93;
  return 100;
}

function normalitzarNomHotel(nom) {
  if (!nom) return "";
  var str = String(nom).trim().toLowerCase();
  
  if (str.indexOf("hilari") !== -1) return "Sant Hilari";
  if (str.indexOf("rivoli") !== -1 || str.indexOf("rambla") !== -1) return "Rivoli Rambla";
  if (str.indexOf("port") !== -1) return "Del Port";
  if (str.indexOf("carlit") !== -1) return "Carlit";
  
  return String(nom).trim();
}

// =========================================================================
// 1. CONSULTA DEL TABLERO DE DECISIÓN (PRICING & REVENUE)
// =========================================================================
function obtenirTaulerDecisioRevenue(hotelSel, dataCapturaRef) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shDades = ss.getSheetByName("Dades_Segments");
    var shNotes = ss.getSheetByName("Estrategia_Notes");
    var shEvents = ss.getSheetByName("Esdeveniments_Mestre");
    var shTanc = ss.getSheetByName("Tancament_LY");

    if (!shDades) return { error: "No s'ha trobat la pestanya 'Dades_Segments'." };

    var dades = shDades.getDataRange().getValues();
    var mapaSnaps = {};
    var mapaHotels = {};
    for (var i = 1; i < dades.length; i++) {
      var h = String(dades[i][0]).trim();
      if (h.toLowerCase().indexOf("arnes") !== -1) continue;
      var s = formatarDataISO(dades[i][1]);
      if (h) mapaHotels[h] = true;
      if (s) mapaSnaps[s] = true;
    }

    var ordrePrioritari = ["Sant Hilari", "Rivoli Rambla", "Del Port", "Carlit"];
    var llistaHotels = [];
    ordrePrioritari.forEach(function(hNom) {
      if (mapaHotels[hNom]) llistaHotels.push(hNom);
    });
    Object.keys(mapaHotels).sort().forEach(function(hNom) {
      if (llistaHotels.indexOf(hNom) === -1) llistaHotels.push(hNom);
    });

    var llistaSnaps = Object.keys(mapaSnaps).sort().reverse();
    if (!hotelSel || llistaHotels.indexOf(hotelSel) === -1) {
      hotelSel = llistaHotels.length > 0 ? llistaHotels[0] : "Sant Hilari";
    }
    if (!dataCapturaRef || llistaSnaps.indexOf(dataCapturaRef) === -1) {
      dataCapturaRef = llistaSnaps.length > 0 ? llistaSnaps[0] : "";
    }

    var capacitatHotel = obtenirCapacitatHotel(hotelSel);

    var mapaTancamentPerData = {};
    if (shTanc) {
      var dTanc = shTanc.getDataRange().getValues();
      var hSelNorm = hotelSel.toLowerCase().trim();
      for (var t = 1; t < dTanc.length; t++) {
        var rowT = dTanc[t];
        var hT = String(rowT[0] || "").toLowerCase().trim();
        if (hT !== hSelNorm) continue;

        var segT = String(rowT[3] || "").toLowerCase().trim();
        if (segT && segT !== "total") continue;

        var dIsoT = formatarDataISO(rowT[1]);
        if (!dIsoT) continue;

        var rnT = Number(rowT[5]) || 0;
        var rRevT = Number(rowT[7]) || 0;
        var fbRevT = Number(rowT[8]) || 0;
        var totRevT = Number(rowT[9]) || (rRevT + fbRevT);

        mapaTancamentPerData[dIsoT] = {
          rn: rnT,
          roomRev: rRevT,
          fbRev: fbRevT,
          totRev: totRevT
        };
      }
    }

    var eventsPerData = {};
    if (shEvents) {
      var dEv = shEvents.getDataRange().getValues();
      var destHotel = (hotelSel.toLowerCase().indexOf("hilari") !== -1) ? "vilar" : "barcelona";
      for (var e = 1; e < dEv.length; e++) {
        var rawDest = String(dEv[e][0] || "").trim().toLowerCase();
        var rawInici = formatarDataISO(dEv[e][1]);
        var rawFi = formatarDataISO(dEv[e][2]) || rawInici;
        var nomEv = String(dEv[e][3] || "").trim();
        if (!nomEv || !rawInici) continue;

        var aplica = (rawDest === "tots" || rawDest === "general" || rawDest === destHotel);
        if (aplica) {
          var dtCur = new Date(rawInici + "T00:00:00");
          var dtMax = new Date(rawFi + "T00:00:00");
          while (dtCur <= dtMax) {
            var isoD = formatarDataISO(dtCur);
            if (!eventsPerData[isoD]) eventsPerData[isoD] = [];
            if (eventsPerData[isoD].indexOf(nomEv) === -1) eventsPerData[isoD].push(nomEv);
            dtCur.setDate(dtCur.getDate() + 1);
          }
        }
      }
    }

    var refDate = dataCapturaRef ? new Date(dataCapturaRef + "T00:00:00") : new Date();
    var anyBase = refDate.getFullYear();
    var mesBase = refDate.getMonth();
    var dataInici = new Date(anyBase, mesBase - 1, 1);
    var dataFi = new Date(anyBase, mesBase + 11, 0);

    var llistaDiesCalendari = [];
    var tempDate = new Date(dataInici.getTime());
    var diesNoms = ["DG", "DL", "DM", "DC", "DJ", "DV", "DS"];

    while (tempDate <= dataFi) {
      var iso = formatarDataISO(tempDate);
      llistaDiesCalendari.push({
        iso: iso,
        mesClau: iso.substring(0, 7),
        diaSetmana: diesNoms[tempDate.getDay()]
      });
      tempDate.setDate(tempDate.getDate() + 1);
    }

    var snapP1 = restarDies(dataCapturaRef, 1);
    var snapP3 = restarDies(dataCapturaRef, 3);
    var snapP7 = restarDies(dataCapturaRef, 7);
    var snapP14 = restarDies(dataCapturaRef, 14);
    var snapP28 = restarDies(dataCapturaRef, 28);

    var notesMap = {};
    if (shNotes) {
      var dNotes = shNotes.getDataRange().getValues();
      for (var n = 1; n < dNotes.length; n++) {
        if (String(dNotes[n][0]).trim() === hotelSel) {
          var dN = formatarDataISO(dNotes[n][1]);
          if (!dN) continue;
          if (!notesMap[dN]) {
            notesMap[dN] = {
              esdeveniment: "", restriccions: "", tarifaAnt: "", tarifaAct: "",
              colorAct: "negre", colorAnt: "negre", grocAct: false, grocAnt: false, estatVenda: "obert"
            };
          }
          var objN = notesMap[dN];
          if (dNotes[n][2]) objN.esdeveniment = dNotes[n][2];
          if (dNotes[n][3]) objN.restriccions = dNotes[n][3];
          if (dNotes[n][4] !== "" && dNotes[n][4] !== null) objN.tarifaAnt = dNotes[n][4];
          if (dNotes[n][5] !== "" && dNotes[n][5] !== null) objN.tarifaAct = dNotes[n][5];
          if (dNotes[n][6] && dNotes[n][6] !== "negre") objN.colorAct = dNotes[n][6];
          if (dNotes[n][7] && dNotes[n][7] !== "negre") objN.colorAnt = dNotes[n][7];
          if (dNotes[n][8] === true || dNotes[n][8] === "true") objN.grocAct = true;
          if (dNotes[n][9] === true || dNotes[n][9] === "true") objN.grocAnt = true;
          
          var estRaw = String(dNotes[n][10] || "").toLowerCase().trim();
          if (estRaw === "true") estRaw = "hotel_tancat";
          if (estRaw) objN.estatVenda = estRaw;
        }
      }
    }

    var mapaCaptures = {};
    for (var r = 1; r < dades.length; r++) {
      var fD = dades[r];
      if (String(fD[0]).trim() !== hotelSel) continue;
      var sn = formatarDataISO(fD[1]);
      var st = formatarDataISO(fD[2]);

      var rnVal = Number(fD[4]) || 0;
      var roomVal = Number(fD[5]) || 0;
      var fbVal = Number(fD[6]) || 0;
      var totVal = Number(fD[7]) || (roomVal + fbVal);

      var rnLYVal = Number(fD[9]) || 0;
      var roomLYVal = Number(fD[10]) || 0;
      var fbLYVal = Number(fD[11]) || 0;
      var totLYVal = Number(fD[12]) || (roomLYVal + fbLYVal);

      if (!mapaCaptures[sn]) mapaCaptures[sn] = {};
      mapaCaptures[sn][st] = {
        rn: rnVal, roomRev: roomVal, fbRev: fbVal, totRev: totVal,
        rnLY: rnLYVal, roomRevLY: roomLYVal, fbRevLY: fbLYVal, totRevLY: totLYVal
      };
    }

    var taulaFinal = [];
    var mesosMap = {};

    llistaDiesCalendari.forEach(function(itemDia) {
      var dStr = itemDia.iso;
      var mClau = itemDia.mesClau;

      var nMan = notesMap[dStr] || { 
        esdeveniment: "", restriccions: "", tarifaAnt: "", tarifaAct: "", 
        colorAct: "negre", colorAnt: "negre", grocAct: false, grocAnt: false, estatVenda: "obert" 
      };

      var evList = eventsPerData[dStr] ? eventsPerData[dStr].slice() : [];
      if (nMan.esdeveniment && evList.indexOf(nMan.esdeveniment) === -1) {
        evList.push(nMan.esdeveniment);
      }

      var notes = {
        esdeveniment: evList.join(" | "),
        restriccions: nMan.restriccions,
        tarifaAnt: nMan.tarifaAnt,
        tarifaAct: nMan.tarifaAct,
        colorAct: nMan.colorAct || "negre",
        colorAnt: nMan.colorAnt || "negre",
        grocAct: !!nMan.grocAct,
        grocAnt: !!nMan.grocAnt,
        estatVenda: nMan.estatVenda || "obert"
      };

      function agafarValors(snapId) {
        if (mapaCaptures[snapId] && mapaCaptures[snapId][dStr]) {
          return { existeix: true, dades: mapaCaptures[snapId][dStr] };
        }
        return { existeix: false, dades: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, rnLY: 0, roomRevLY: 0, fbRevLY: 0, totRevLY: 0 } };
      }

      var cActual = agafarValors(dataCapturaRef).dades;
      var cP1 = agafarValors(snapP1);
      var cP3 = agafarValors(snapP3);
      var cP7 = agafarValors(snapP7);
      var cP14 = agafarValors(snapP14);
      var cP28 = agafarValors(snapP28);

      function calcBloc(rn, room, fb, tot) {
        var occ = capacitatHotel > 0 ? (rn / capacitatHotel) * 100 : 0;
        var adr = rn > 0 ? (room / rn) : 0;
        return { occ: occ, rn: rn, adr: adr, roomRev: room, fbRev: fb, totRev: tot };
      }

      function calcPickup(base, anteriorObj) {
        if (!anteriorObj.existeix) return null;
        var ant = anteriorObj.dades;
        var dRN = base.rn - ant.rn;
        var dRoom = base.roomRev - ant.roomRev;
        var dFB = base.fbRev - ant.fbRev;
        var dTot = base.totRev - ant.totRev;

        var adrBase = base.rn > 0 ? (base.roomRev / base.rn) : 0;
        var adrAnt = ant.rn > 0 ? (ant.roomRev / ant.rn) : 0;
        var dADR = (base.rn > 0 && ant.rn > 0) ? (adrBase - adrAnt) : (dRN !== 0 ? (dRoom / dRN) : 0);
        var adrNou = dRN > 0 ? (dRoom / dRN) : 0;

        return { rn: dRN, adrDiff: dADR, adrNou: adrNou, roomRev: dRoom, fbRev: dFB, totRev: dTot };
      }

      function calcPickupLY(base, anteriorObj) {
        if (!anteriorObj.existeix) return null;
        var ant = anteriorObj.dades;
        var dRN = base.rnLY - ant.rnLY;
        var dRoom = base.roomRevLY - ant.roomRevLY;
        var dFB = base.fbRevLY - ant.fbRevLY;
        var dTot = base.totRevLY - ant.totRevLY;

        var adrBaseLY = base.rnLY > 0 ? (base.roomRevLY / base.rnLY) : 0;
        var adrAntLY = ant.rnLY > 0 ? (ant.roomRevLY / ant.rnLY) : 0;
        var dADR = (base.rnLY > 0 && ant.rnLY > 0) ? (adrBaseLY - adrAntLY) : (dRN !== 0 ? (dRoom / dRN) : 0);
        var adrNouLY = dRN > 0 ? (dRoom / dRN) : 0;

        return { rn: dRN, adrDiff: dADR, adrNou: adrNouLY, roomRev: dRoom, fbRev: dFB, totRev: dTot };
      }

      var oBloc = calcBloc(cActual.rn, cActual.roomRev, cActual.fbRev, cActual.totRev);
      var sBloc = calcBloc(cActual.rnLY, cActual.roomRevLY, cActual.fbRevLY, cActual.totRevLY);

      var dataAnyAnterior = obtenirDataNaturalAnyAnterior(dStr);
      var tancDades = mapaTancamentPerData[dataAnyAnterior] || { rn: 0, roomRev: 0, fbRev: 0, totRev: 0 };
      var tancBloc = calcBloc(tancDades.rn, tancDades.roomRev, tancDades.fbRev, tancDades.totRev);

      var p1 = calcPickup(cActual, cP1);
      var p3 = calcPickup(cActual, cP3);
      var p7 = calcPickup(cActual, cP7);
      var p14 = calcPickup(cActual, cP14);
      var p28 = calcPickup(cActual, cP28);

      var pLY1 = calcPickupLY(cActual, cP1);
      var pLY3 = calcPickupLY(cActual, cP3);
      var pLY7 = calcPickupLY(cActual, cP7);
      var pLY14 = calcPickupLY(cActual, cP14);
      var pLY28 = calcPickupLY(cActual, cP28);

      if (!mesosMap[mClau]) {
        mesosMap[mClau] = {
          capacitatAcumulada: 0,
          otb: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0 },
          stly: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0 },
          tancament: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0 },
          p1: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          p3: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          p7: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          p14: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          p28: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          pLY1: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          pLY3: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          pLY7: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          pLY14: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false },
          pLY28: { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, valid: false }
        };
      }

      var mObj = mesosMap[mClau];
      mObj.capacitatAcumulada += capacitatHotel;
      mObj.otb.rn += oBloc.rn; mObj.otb.roomRev += oBloc.roomRev; mObj.otb.fbRev += oBloc.fbRev; mObj.otb.totRev += oBloc.totRev;
      mObj.stly.rn += sBloc.rn; mObj.stly.roomRev += sBloc.roomRev; mObj.stly.fbRev += sBloc.fbRev; mObj.stly.totRev += sBloc.totRev;
      mObj.tancament.rn += tancBloc.rn; mObj.tancament.roomRev += tancBloc.roomRev; mObj.tancament.fbRev += tancBloc.fbRev; mObj.tancament.totRev += tancBloc.totRev;

      function sumaPK(acc, item) {
        if (!item) return;
        acc.valid = true;
        acc.rn += item.rn; acc.roomRev += item.roomRev; acc.fbRev += item.fbRev; acc.totRev += item.totRev;
      }
      sumaPK(mObj.p1, p1); sumaPK(mObj.p3, p3); sumaPK(mObj.p7, p7); sumaPK(mObj.p14, p14); sumaPK(mObj.p28, p28);
      sumaPK(mObj.pLY1, pLY1); sumaPK(mObj.pLY3, pLY3); sumaPK(mObj.pLY7, pLY7); sumaPK(mObj.pLY14, pLY14); sumaPK(mObj.pLY28, pLY28);

      taulaFinal.push({
        esTotalMes: false,
        mesClau: mClau,
        data: dStr,
        dia: itemDia.diaSetmana,
        notes: notes,
        capacitatHotel: capacitatHotel,
        otb: oBloc,
        pk1: p1, pk3: p3, pk7: p7, pk14: p14, pk28: p28,
        stly: sBloc,
        pkLY1: pLY1, pkLY3: pLY3, pkLY7: pLY7, pkLY14: pLY14, pkLY28: pLY28,
        tancament: tancBloc
      });
    });

    var llistaAmbTotals = [];
    var nomsMesos = ["Gener", "Febrer", "Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre", "Novembre", "Desembre"];

    for (var i = 0; i < taulaFinal.length; i++) {
      llistaAmbTotals.push(taulaFinal[i]);
      var diaActual = taulaFinal[i];
      var diaSeguent = taulaFinal[i + 1];

      if (!diaSeguent || diaSeguent.mesClau !== diaActual.mesClau) {
        var mK = diaActual.mesClau;
        var totM = mesosMap[mK];
        var pParts = mK.split("-");
        var numMes = parseInt(pParts[1], 10) - 1;
        var nomMesComplet = nomsMesos[numMes] + " " + pParts[0];

        function tBloc(acc) {
          var occ = totM.capacitatAcumulada > 0 ? (acc.rn / totM.capacitatAcumulada) * 100 : 0;
          var adr = acc.rn > 0 ? (acc.roomRev / acc.rn) : 0;
          return { occ: occ, rn: acc.rn, adr: adr, roomRev: acc.roomRev, fbRev: acc.fbRev, totRev: acc.totRev };
        }

        function tPK(acc) {
          if (!acc.valid) return null;
          var adr = acc.rn !== 0 ? (acc.roomRev / acc.rn) : 0;
          return { rn: acc.rn, adrDiff: adr, adrNou: adr, roomRev: acc.roomRev, fbRev: acc.fbRev, totRev: acc.totRev };
        }

        llistaAmbTotals.push({
          esTotalMes: true,
          mesClau: mK,
          data: "TOTAL " + nomMesComplet.toUpperCase(),
          dia: "",
          notes: { esdeveniment: "", restriccions: "", tarifaAnt: "", tarifaAct: "", colorAct: "negre", colorAnt: "negre", grocAct: false, grocAnt: false, estatVenda: "obert" },
          capacitatHotel: totM.capacitatAcumulada,
          otb: tBloc(totM.otb),
          pk1: tPK(totM.p1), pk3: tPK(totM.p3), pk7: tPK(totM.p7), pk14: tPK(totM.p14), pk28: tPK(totM.p28),
          stly: tBloc(totM.stly),
          pkLY1: tPK(totM.pLY1), pkLY3: tPK(totM.pLY3), pkLY7: tPK(totM.pLY7), pkLY14: tPK(totM.pLY14), pkLY28: tPK(totM.pLY28),
          tancament: tBloc(totM.tancament)
        });
      }
    }

    return {
      hotelActual: hotelSel,
      capacitat: capacitatHotel,
      llistaHotels: llistaHotels,
      llistaSnaps: llistaSnaps,
      dataCapturaRef: dataCapturaRef,
      potEditarPricing: verificarPermisEdicioPricing(),
      filas: llistaAmbTotals
    };
  } catch (err) {
    return { error: "Error en el servidor: " + err.message };
  }
}

// =========================================================================
// 2. CONSULTA DE FONTS DE NEGOCI
// =========================================================================
function obtenirTaulerFontsNegoci(hotelSel, dataCapturaRef) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shFN = ss.getSheetByName("Dades_Fonts_Negoci");
    if (!shFN) return { error: "No s'ha trobat la pestanya 'Dades_Fonts_Negoci'." };

    var dadesFN = shFN.getDataRange().getValues();
    var mapaSnaps = {};
    var mapaHotels = {};

    for (var i = 1; i < dadesFN.length; i++) {
      var h = String(dadesFN[i][0]).trim();
      if (h.toLowerCase().indexOf("arnes") !== -1) continue;
      var s = formatarDataISO(dadesFN[i][1]);
      if (h) mapaHotels[h] = true;
      if (s) mapaSnaps[s] = true;
    }

    var ordrePrioritari = ["Sant Hilari", "Rivoli Rambla", "Del Port", "Carlit"];
    var llistaHotels = [];
    ordrePrioritari.forEach(function(hNom) {
      if (mapaHotels[hNom]) llistaHotels.push(hNom);
    });
    Object.keys(mapaHotels).sort().forEach(function(hNom) {
      if (llistaHotels.indexOf(hNom) === -1) llistaHotels.push(hNom);
    });

    var llistaSnaps = Object.keys(mapaSnaps).sort().reverse();
    if (!hotelSel || llistaHotels.indexOf(hotelSel) === -1) {
      hotelSel = llistaHotels.length > 0 ? llistaHotels[0] : "Sant Hilari";
    }
    if (!dataCapturaRef || llistaSnaps.indexOf(dataCapturaRef) === -1) {
      dataCapturaRef = llistaSnaps.length > 0 ? llistaSnaps[0] : "";
    }

    var fontsLlista = [
      "Web Propia", "Particulars", "Agencies Online", "Agencies Online Net",
      "Altres Agencies", "Empreses Grups", "Empreses Agencia", "Col.lectius"
    ];

    var snapP1 = restarDies(dataCapturaRef, 1);
    var snapP7 = restarDies(dataCapturaRef, 7);

    var mapaMatriu = {};
    for (var r = 1; r < dadesFN.length; r++) {
      var fD = dadesFN[r];
      if (String(fD[0]).trim() !== hotelSel) continue;
      var sn = formatarDataISO(fD[1]);
      var st = formatarDataISO(fD[2]);
      var canal = String(fD[3] || "").trim();

      if (!mapaMatriu[sn]) mapaMatriu[sn] = {};
      if (!mapaMatriu[sn][st]) mapaMatriu[sn][st] = { fonts: {}, total: { rn: 0, roomRev: 0, rnLY: 0, roomRevLY: 0 } };

      var rn = Number(fD[4]) || 0;
      var rev = Number(fD[5]) || 0;
      var rnLY = Number(fD[6]) || 0;
      var revLY = Number(fD[7]) || 0;

      mapaMatriu[sn][st].fonts[canal] = { rn: rn, roomRev: rev, rnLY: rnLY, roomRevLY: revLY };
      mapaMatriu[sn][st].total.rn += rn;
      mapaMatriu[sn][st].total.roomRev += rev;
      mapaMatriu[sn][st].total.rnLY += rnLY;
      mapaMatriu[sn][st].total.roomRevLY += revLY;
    }

    var refDate = dataCapturaRef ? new Date(dataCapturaRef + "T00:00:00") : new Date();
    var anyBase = refDate.getFullYear();
    var mesBase = refDate.getMonth();
    var dataInici = new Date(anyBase, mesBase - 1, 1);
    var dataFi = new Date(anyBase, mesBase + 11, 0);

    var llistaDies = [];
    var tempDate = new Date(dataInici.getTime());
    var diesNoms = ["DG", "DL", "DM", "DC", "DJ", "DV", "DS"];

    while (tempDate <= dataFi) {
      var iso = formatarDataISO(tempDate);
      llistaDies.push({
        iso: iso,
        mesClau: iso.substring(0, 7),
        diaSetmana: diesNoms[tempDate.getDay()]
      });
      tempDate.setDate(tempDate.getDate() + 1);
    }

    var filesRetorn = [];
    var totalsMes = {};

    llistaDies.forEach(function(itemDia) {
      var dStr = itemDia.iso;
      var mK = itemDia.mesClau;

      function obtenirObj(snapId) {
        if (mapaMatriu[snapId] && mapaMatriu[snapId][dStr]) return mapaMatriu[snapId][dStr];
        return { fonts: {}, total: { rn: 0, roomRev: 0, rnLY: 0, roomRevLY: 0 } };
      }

      var act = obtenirObj(dataCapturaRef);
      var p1 = obtenirObj(snapP1);
      var p7 = obtenirObj(snapP7);

      function calcMetriques(rn, rev) {
        return { rn: rn, roomRev: rev, adr: rn > 0 ? (rev / rn) : 0 };
      }

      function calcPickupObj(bAct, bAnt) {
        var res = { fonts: {}, total: null };
        var dRN_tot = bAct.total.rn - bAnt.total.rn;
        var dRev_tot = bAct.total.roomRev - bAnt.total.roomRev;
        var adrBaseTot = bAct.total.rn > 0 ? (bAct.total.roomRev / bAct.total.rn) : 0;
        var adrAntTot = bAnt.total.rn > 0 ? (bAnt.total.roomRev / bAnt.total.rn) : 0;
        var dADR_tot = (bAct.total.rn > 0 && bAnt.total.rn > 0) ? (adrBaseTot - adrAntTot) : (dRN_tot !== 0 ? (dRev_tot / dRN_tot) : 0);
        res.total = { rn: dRN_tot, roomRev: dRev_tot, adr: dADR_tot };

        fontsLlista.forEach(function(fn) {
          var fa = bAct.fonts[fn] || { rn: 0, roomRev: 0 };
          var fp = bAnt.fonts[fn] || { rn: 0, roomRev: 0 };
          var dRN = fa.rn - fp.rn;
          var dRev = fa.roomRev - fp.roomRev;
          var adrBase = fa.rn > 0 ? (fa.roomRev / fa.rn) : 0;
          var adrAnt = fp.rn > 0 ? (fp.roomRev / fp.rn) : 0;
          var dADR = (fa.rn > 0 && fp.rn > 0) ? (adrBase - adrAnt) : (dRN !== 0 ? (dRev / dRN) : 0);
          res.fonts[fn] = { rn: dRN, roomRev: dRev, adr: dADR };
        });
        return res;
      }

      var objOTB = { fonts: {}, total: calcMetriques(act.total.rn, act.total.roomRev) };
      var objSTLY = { fonts: {}, total: calcMetriques(act.total.rnLY, act.total.roomRevLY) };

      fontsLlista.forEach(function(fn) {
        var fa = act.fonts[fn] || { rn: 0, roomRev: 0, rnLY: 0, roomRevLY: 0 };
        objOTB.fonts[fn] = calcMetriques(fa.rn, fa.roomRev);
        objSTLY.fonts[fn] = calcMetriques(fa.rnLY, fa.roomRevLY);
      });

      var objP1 = calcPickupObj(act, p1);
      var objP7 = calcPickupObj(act, p7);

      if (!totalsMes[mK]) {
        totalsMes[mK] = {
          otb: { fonts: {}, total: { rn: 0, roomRev: 0 } },
          p1: { fonts: {}, total: { rn: 0, roomRev: 0 } },
          p7: { fonts: {}, total: { rn: 0, roomRev: 0 } },
          stly: { fonts: {}, total: { rn: 0, roomRev: 0 } }
        };
      }

      function acumular(acc, obj) {
        acc.total.rn += obj.total.rn;
        acc.total.roomRev += obj.total.roomRev;
        fontsLlista.forEach(function(fn) {
          if (!acc.fonts[fn]) acc.fonts[fn] = { rn: 0, roomRev: 0 };
          var fo = obj.fonts[fn] || { rn: 0, roomRev: 0 };
          acc.fonts[fn].rn += fo.rn;
          acc.fonts[fn].roomRev += fo.roomRev;
        });
      }

      acumular(totalsMes[mK].otb, objOTB);
      acumular(totalsMes[mK].p1, objP1);
      acumular(totalsMes[mK].p7, objP7);
      acumular(totalsMes[mK].stly, objSTLY);

      filesRetorn.push({
        data: dStr,
        dia: itemDia.diaSetmana,
        mesClau: mK,
        otb: objOTB,
        p1: objP1,
        p7: objP7,
        stly: objSTLY
      });
    });

    return {
      hotelActual: hotelSel,
      llistaHotels: llistaHotels,
      llistaSnaps: llistaSnaps,
      dataCapturaRef: dataCapturaRef,
      fontsLlista: fontsLlista,
      filas: filesRetorn,
      totalsMes: totalsMes
    };
  } catch (err) {
    return { error: "Error Fonts: " + err.message };
  }
}

// =========================================================================
// 3. CONSULTA DE TIPOLOGIES D'HABITACIÓ (AMB NOMS COMERCIALS I INVENTARI)
// =========================================================================
function obtenirConfigTipologiesHotel(hotelSel) {
  var h = normalitzarNomHotel(hotelSel);
  var configs = {
    "Rivoli Rambla": [
      { clau: "Confort",               nom: "Deluxe",          inv: 36 },
      { clau: "Confort Vistes/Jardí",  nom: "DLX V. Rambla",   inv: 21 },
      { clau: "Familiar",              nom: "DLX Triple",      inv: 12 },
      { clau: "Confort especial",      nom: "Deluxe+",         inv: 17 },
      { clau: "Familiar especial",     nom: "DLX V. Rambla+",  inv: 10 },
      { clau: "Superior",              nom: "Superior",        inv: 9  },
      { clau: "Suite",                 nom: "Suite",           inv: 1  },
      { clau: "Suite Exclusiva",       nom: "Suite Exclusive", inv: 4  },
      { clau: "Economica",             nom: "Small",           inv: 16 }
    ],
    "Del Port": [
      { clau: "Confort",               nom: "Confort",         inv: 30 },
      { clau: "Confort Vistes/Jardí",  nom: "Confort V.",      inv: 6  },
      { clau: "Familiar Vistes/Jardí", nom: "Triple",          inv: 4  },
      { clau: "Economica",             nom: "Essential",       inv: 8  }
    ],
    "Carlit": [
      { clau: "Confort",               nom: "Confort",         inv: 19 },
      { clau: "Confort Vistes/Jardí",  nom: "Confort B.",      inv: 4  },
      { clau: "Familiar",              nom: "Triple",          inv: 5  },
      { clau: "Confort especial",      nom: "Standard Ind.",   inv: 5  },
      { clau: "Economica",             nom: "Small",           inv: 5  }
    ],
    "Sant Hilari": [
      { clau: "Confort",               nom: "Confort TRP.",    inv: 52 },
      { clau: "Familiar",              nom: "Confort QUAD.",   inv: 7  },
      { clau: "Familiar Vistes/Jardí", nom: "Jardí",           inv: 26 },
      { clau: "Confort especial",      nom: "Baby room",       inv: 6  },
      { clau: "Superior",              nom: "Baby Suite",      inv: 1  },
      { clau: "Suite",                 nom: "Suite",           inv: 1  }
    ]
  };
  return configs[h] || [];
}

function obtenirTaulerTipologies(hotelSel, dataCapturaRef) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var shTP = ss.getSheetByName("Dades_Tipologies");
    if (!shTP) return { error: "No s'ha trobat la pestanya 'Dades_Tipologies'." };

    var dadesTP = shTP.getDataRange().getValues();
    var mapaSnaps = {};
    var mapaHotels = {};

    for (var i = 1; i < dadesTP.length; i++) {
      var h = String(dadesTP[i][0]).trim();
      if (h.toLowerCase().indexOf("arnes") !== -1) continue;
      var s = formatarDataISO(dadesTP[i][1]);
      if (h) mapaHotels[h] = true;
      if (s) mapaSnaps[s] = true;
    }

    var ordrePrioritari = ["Sant Hilari", "Rivoli Rambla", "Del Port", "Carlit"];
    var llistaHotels = [];
    ordrePrioritari.forEach(function(hNom) {
      if (mapaHotels[hNom]) llistaHotels.push(hNom);
    });
    Object.keys(mapaHotels).sort().forEach(function(hNom) {
      if (llistaHotels.indexOf(hNom) === -1) llistaHotels.push(hNom);
    });

    var llistaSnaps = Object.keys(mapaSnaps).sort().reverse();
    if (!hotelSel || llistaHotels.indexOf(hotelSel) === -1) {
      hotelSel = llistaHotels.length > 0 ? llistaHotels[0] : "Sant Hilari";
    }
    if (!dataCapturaRef || llistaSnaps.indexOf(dataCapturaRef) === -1) {
      dataCapturaRef = llistaSnaps.length > 0 ? llistaSnaps[0] : "";
    }

    var snapP1 = restarDies(dataCapturaRef, 1);
    var snapP7 = restarDies(dataCapturaRef, 7);

    var configHotel = obtenirConfigTipologiesHotel(hotelSel);
    var tipusActius = [];
    var diccionariNoms = {};
    var inventariTipus = {};

    configHotel.forEach(function(item) {
      tipusActius.push(item.clau);
      diccionariNoms[item.clau] = item.nom + " (" + item.inv + ")";
      inventariTipus[item.clau] = item.inv;
    });

    var mapaMatriu = {};
    for (var r = 1; r < dadesTP.length; r++) {
      var fD = dadesTP[r];
      if (String(fD[0]).trim() !== hotelSel) continue;
      var sn = formatarDataISO(fD[1]);
      var st = formatarDataISO(fD[2]);
      var tipus = String(fD[3] || "").trim();

      if (!mapaMatriu[sn]) mapaMatriu[sn] = {};
      if (!mapaMatriu[sn][st]) mapaMatriu[sn][st] = { tipus: {}, total: { rn: 0, roomRev: 0, rnLY: 0, roomRevLY: 0 } };

      var rn = Number(fD[4]) || 0;
      var rev = Number(fD[5]) || 0;
      var rnLY = Number(fD[9]) || 0;
      var revLY = Number(fD[10]) || 0;

      mapaMatriu[sn][st].tipus[tipus] = { rn: rn, roomRev: rev, rnLY: rnLY, roomRevLY: revLY };
      mapaMatriu[sn][st].total.rn += rn;
      mapaMatriu[sn][st].total.roomRev += rev;
      mapaMatriu[sn][st].total.rnLY += rnLY;
      mapaMatriu[sn][st].total.roomRevLY += revLY;
    }

    var refDate = dataCapturaRef ? new Date(dataCapturaRef + "T00:00:00") : new Date();
    var anyBase = refDate.getFullYear();
    var mesBase = refDate.getMonth();
    var dataInici = new Date(anyBase, mesBase - 1, 1);
    var dataFi = new Date(anyBase, mesBase + 11, 0);

    var llistaDies = [];
    var tempDate = new Date(dataInici.getTime());
    var diesNoms = ["DG", "DL", "DM", "DC", "DJ", "DV", "DS"];

    while (tempDate <= dataFi) {
      var iso = formatarDataISO(tempDate);
      llistaDies.push({
        iso: iso,
        mesClau: iso.substring(0, 7),
        diaSetmana: diesNoms[tempDate.getDay()]
      });
      tempDate.setDate(tempDate.getDate() + 1);
    }

    var filesRetorn = [];
    var totalsMes = {};

    llistaDies.forEach(function(itemDia) {
      var dStr = itemDia.iso;
      var mK = itemDia.mesClau;

      function obtenirObj(snapId) {
        if (mapaMatriu[snapId] && mapaMatriu[snapId][dStr]) return mapaMatriu[snapId][dStr];
        return { tipus: {}, total: { rn: 0, roomRev: 0, rnLY: 0, roomRevLY: 0 } };
      }

      var act = obtenirObj(dataCapturaRef);
      var p1 = obtenirObj(snapP1);
      var p7 = obtenirObj(snapP7);

      function calcMetriques(rn, rev) {
        return { rn: rn, roomRev: rev, adr: rn > 0 ? (rev / rn) : 0 };
      }

      function calcPickupObj(bAct, bAnt) {
        var res = { tipus: {}, total: null };
        var dRN_tot = bAct.total.rn - bAnt.total.rn;
        var dRev_tot = bAct.total.roomRev - bAnt.total.roomRev;
        var adrBaseTot = bAct.total.rn > 0 ? (bAct.total.roomRev / bAct.total.rn) : 0;
        var adrAntTot = bAnt.total.rn > 0 ? (bAnt.total.roomRev / bAnt.total.rn) : 0;
        var dADR_tot = (bAct.total.rn > 0 && bAnt.total.rn > 0) ? (adrBaseTot - adrAntTot) : (dRN_tot !== 0 ? (dRev_tot / dRN_tot) : 0);
        res.total = { rn: dRN_tot, roomRev: dRev_tot, adr: dADR_tot };

        tipusActius.forEach(function(tp) {
          var ta = bAct.tipus[tp] || { rn: 0, roomRev: 0 };
          var tpAnt = bAnt.tipus[tp] || { rn: 0, roomRev: 0 };
          var dRN = ta.rn - tpAnt.rn;
          var dRev = ta.roomRev - tpAnt.roomRev;
          var adrBase = ta.rn > 0 ? (ta.roomRev / ta.rn) : 0;
          var adrAnt = tpAnt.rn > 0 ? (tpAnt.roomRev / tpAnt.rn) : 0;
          var dADR = (ta.rn > 0 && tpAnt.rn > 0) ? (adrBase - adrAnt) : (dRN !== 0 ? (dRev / dRN) : 0);
          res.tipus[tp] = { rn: dRN, roomRev: dRev, adr: dADR };
        });
        return res;
      }

      var objOTB = { tipus: {}, total: calcMetriques(act.total.rn, act.total.roomRev) };
      var objSTLY = { tipus: {}, total: calcMetriques(act.total.rnLY, act.total.roomRevLY) };

      tipusActius.forEach(function(tp) {
        var ta = act.tipus[tp] || { rn: 0, roomRev: 0, rnLY: 0, roomRevLY: 0 };
        objOTB.tipus[tp] = calcMetriques(ta.rn, ta.roomRev);
        objSTLY.tipus[tp] = calcMetriques(ta.rnLY, ta.roomRevLY);
      });

      var objP1 = calcPickupObj(act, p1);
      var objP7 = calcPickupObj(act, p7);

      if (!totalsMes[mK]) {
        totalsMes[mK] = {
          otb: { tipus: {}, total: { rn: 0, roomRev: 0 } },
          p1: { tipus: {}, total: { rn: 0, roomRev: 0 } },
          p7: { tipus: {}, total: { rn: 0, roomRev: 0 } },
          stly: { tipus: {}, total: { rn: 0, roomRev: 0 } }
        };
      }

      function acumular(acc, obj) {
        acc.total.rn += obj.total.rn;
        acc.total.roomRev += obj.total.roomRev;
        tipusActius.forEach(function(tp) {
          if (!acc.tipus[tp]) acc.tipus[tp] = { rn: 0, roomRev: 0 };
          var tObj = obj.tipus[tp] || { rn: 0, roomRev: 0 };
          acc.tipus[tp].rn += tObj.rn;
          acc.tipus[tp].roomRev += tObj.roomRev;
        });
      }

      acumular(totalsMes[mK].otb, objOTB);
      acumular(totalsMes[mK].p1, objP1);
      acumular(totalsMes[mK].p7, objP7);
      acumular(totalsMes[mK].stly, objSTLY);

      filesRetorn.push({
        data: dStr,
        dia: itemDia.diaSetmana,
        mesClau: mK,
        otb: objOTB,
        p1: objP1,
        p7: objP7,
        stly: objSTLY
      });
    });

    return {
      hotelActual: hotelSel,
      llistaHotels: llistaHotels,
      llistaSnaps: llistaSnaps,
      dataCapturaRef: dataCapturaRef,
      tipusActius: tipusActius,
      diccionariNoms: diccionariNoms,
      inventariTipus: inventariTipus,
      filas: filesRetorn,
      totalsMes: totalsMes
    };
  } catch (err) {
    return { error: "Error Tipologies: " + err.message };
  }
}

// =========================================================================
// 4. FITXA MENSUAL EXECUTIVA (DEFINITIVA I CALIBRADA)
// =========================================================================
function obtenirFitxaMensualExecutiva(hotelSel, dataCapturaRef, mesClauSel) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet() || (typeof SPREADSHEET_ID !== 'undefined' ? SpreadsheetApp.openById(SPREADSHEET_ID) : null);
    if (!ss) return { error: "No s'ha pogut accedir al full de càlcul." };

    var shSeg = ss.getSheetByName("Dades_Segments");
    var shFN = ss.getSheetByName("Dades_Fonts_Negoci");
    var shPress = ss.getSheetByName("Pressupost");
    var shTanc = ss.getSheetByName("Tancament_LY");

    if (!shSeg) return { error: "No s'ha trobat la pestanya 'Dades_Segments'." };

    var llistaHotels = ["Sant Hilari", "Rivoli Rambla", "Del Port", "Carlit"];
    if (!hotelSel || llistaHotels.indexOf(hotelSel) === -1) {
      hotelSel = llistaHotels[0];
    }

    var hotelSelNorm = normalitzarNomHotel(hotelSel).toLowerCase();
    var capacitatHotel = obtenirCapacitatHotel(hotelSel);
    var dadesSeg = shSeg.getDataRange().getValues();

    function dataYMD(v) {
      if (!v) return "";
      if (v instanceof Date) {
        var y = v.getFullYear();
        var m = v.getMonth() + 1;
        var d = v.getDate();
        return y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);
      }
      var s = String(v).trim().split("T")[0];
      var p = s.split(/[\/\-]/);
      if (p.length === 3) {
        if (p[0].length === 4) return p[0] + "-" + (p[1].length === 1 ? "0" + p[1] : p[1]) + "-" + (p[2].length === 1 ? "0" + p[2] : p[2]);
        var a = p[2].length === 2 ? "20" + p[2] : p[2];
        return a + "-" + (p[1].length === 1 ? "0" + p[1] : p[1]) + "-" + (p[0].length === 1 ? "0" + p[0] : p[0]);
      }
      return s.substring(0, 10);
    }

    var mapaSnapsSeg = {};
    for (var i = 1; i < dadesSeg.length; i++) {
      if (normalitzarNomHotel(dadesSeg[i][0]).toLowerCase() === hotelSelNorm) {
        var sIso = dataYMD(dadesSeg[i][1]);
        if (sIso) mapaSnapsSeg[sIso] = true;
      }
    }
    var llistaSnaps = Object.keys(mapaSnapsSeg).sort().reverse();
    if (!dataCapturaRef || llistaSnaps.indexOf(dataYMD(dataCapturaRef)) === -1) {
      dataCapturaRef = llistaSnaps.length > 0 ? llistaSnaps[0] : dataYMD(new Date());
    } else {
      dataCapturaRef = dataYMD(dataCapturaRef);
    }

    function trobarSnapEnLlista(refYMD, diesObj, llista) {
      if (!refYMD || !llista || llista.length === 0) return "";
      var tRef = new Date(refYMD + "T00:00:00").getTime();
      var tObj = tRef - (diesObj * 86400000);
      var millor = "";
      var minDiff = Infinity;
      for (var k = 0; k < llista.length; k++) {
        var sn = llista[k];
        var tSn = new Date(sn + "T00:00:00").getTime();
        if (tSn < tRef) {
          var diff = Math.abs(tSn - tObj);
          if (diff < minDiff) {
            minDiff = diff;
            millor = sn;
          }
        }
      }
      return millor;
    }

    var snapP7_Seg = trobarSnapEnLlista(dataCapturaRef, 7, llistaSnaps);

    var refDate = new Date(dataCapturaRef + "T00:00:00");
    var anyBase = refDate.getFullYear();
    var mesBase = refDate.getMonth();

    var llistaMesos = [];
    var nomsMesos = ["Gener", "Febrer", "Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre", "Novembre", "Desembre"];

    for (var mOffset = 0; mOffset < 12; mOffset++) {
      var dIter = new Date(anyBase, mesBase + mOffset, 1);
      var cAny = dIter.getFullYear();
      var cM = dIter.getMonth() + 1;
      var cClau = cAny + "-" + (cM < 10 ? "0" + cM : cM);
      llistaMesos.push({ clau: cClau, nom: nomsMesos[cM - 1] + " " + cAny });
    }

    if (!mesClauSel) {
      mesClauSel = llistaMesos.length > 0 ? llistaMesos[0].clau : (anyBase + "-" + ((mesBase + 1) < 10 ? "0" + (mesBase + 1) : (mesBase + 1)));
    }

    var partsM = mesClauSel.split("-");
    var anySel = parseInt(partsM[0], 10);
    var mesIdx = parseInt(partsM[1], 10) - 1;
    var diesAlMes = new Date(anySel, mesIdx + 1, 0).getDate();
    var capacitatMes = capacitatHotel * diesAlMes;
    var nomMesComplet = nomsMesos[mesIdx] + " " + anySel;
    var esMesEnCurs = (mesClauSel === dataCapturaRef.substring(0, 7));

    var kpis = {
      avui: { occ: 0, rn: 0, adr: 0, revpar: 0, roomRev: 0, fbRev: 0, totRev: 0 },
      ly: { occ: 0, rn: 0, adr: 0, revpar: 0, roomRev: 0, fbRev: 0, totRev: 0 },
      var7d: { occ: 0, rn: 0, adr: 0, revpar: 0, roomRev: 0, fbRev: 0, totRev: 0 },
      var7dLY: { occ: 0, rn: 0, adr: 0, revpar: 0, roomRev: 0, fbRev: 0, totRev: 0 },
      pressupost: { occ: 0, rn: 0, adr: 0, revpar: 0, roomRev: 0, fbRev: 0, totRev: 0 },
      varPress: { occ: 0, rn: 0, adr: 0, revpar: 0, roomRev: 0, fbRev: 0, totRev: 0 },
      tancament: { occ: 0, rn: 0, adr: 0, revpar: 0, roomRev: 0, fbRev: 0, totRev: 0 }
    };

    var acumulatsP7 = { rn: 0, roomRev: 0, fbRev: 0, totRev: 0, rnLY: 0, roomRevLY: 0, fbRevLY: 0, totRevLY: 0 };

    for (var r = 1; r < dadesSeg.length; r++) {
      var row = dadesSeg[r];
      if (normalitzarNomHotel(row[0]).toLowerCase() !== hotelSelNorm) continue;

      var estadaD = dataYMD(row[2]);
      if (estadaD.indexOf(mesClauSel) !== 0) continue;

      var snapD = dataYMD(row[1]);
      var rn = Number(row[4]) || 0;
      var room = Number(row[5]) || 0;
      var fb = Number(row[6]) || 0;
      var tot = Number(row[7]) || (room + fb);

      var rnLY = Number(row[9]) || 0;
      var roomLY = Number(row[10]) || 0;
      var fbLY = Number(row[11]) || 0;
      var totLY = Number(row[12]) || (roomLY + fbLY);

      if (snapD === dataCapturaRef) {
        kpis.avui.rn += rn;
        kpis.avui.roomRev += room;
        kpis.avui.fbRev += fb;
        kpis.avui.totRev += tot;

        kpis.ly.rn += rnLY;
        kpis.ly.roomRev += roomLY;
        kpis.ly.fbRev += fbLY;
        kpis.ly.totRev += totLY;
      } else if (snapD === snapP7_Seg) {
        acumulatsP7.rn += rn;
        acumulatsP7.roomRev += room;
        acumulatsP7.fbRev += fb;
        acumulatsP7.totRev += tot;

        acumulatsP7.rnLY += rnLY;
        acumulatsP7.roomRevLY += roomLY;
        acumulatsP7.fbRevLY += fbLY;
        acumulatsP7.totRevLY += totLY;
      }
    }

    kpis.avui.occ = capacitatMes > 0 ? (kpis.avui.rn / capacitatMes) * 100 : 0;
    kpis.avui.adr = kpis.avui.rn > 0 ? (kpis.avui.roomRev / kpis.avui.rn) : 0;
    kpis.avui.revpar = capacitatMes > 0 ? (kpis.avui.roomRev / capacitatMes) : 0;

    kpis.ly.occ = capacitatMes > 0 ? (kpis.ly.rn / capacitatMes) * 100 : 0;
    kpis.ly.adr = kpis.ly.rn > 0 ? (kpis.ly.roomRev / kpis.ly.rn) : 0;
    kpis.ly.revpar = capacitatMes > 0 ? (kpis.ly.roomRev / capacitatMes) : 0;

    var diffRN_7D = kpis.avui.rn - acumulatsP7.rn;
    var diffRoom_7D = kpis.avui.roomRev - acumulatsP7.roomRev;
    kpis.var7d.rn = diffRN_7D;
    kpis.var7d.roomRev = diffRoom_7D;
    kpis.var7d.fbRev = kpis.avui.fbRev - acumulatsP7.fbRev;
    kpis.var7d.totRev = kpis.avui.totRev - acumulatsP7.totRev;
    kpis.var7d.occ = capacitatMes > 0 ? (diffRN_7D / capacitatMes) * 100 : 0;
    kpis.var7d.adr = (kpis.avui.rn > 0 && acumulatsP7.rn > 0) ? (kpis.avui.adr - (acumulatsP7.roomRev / acumulatsP7.rn)) : (diffRN_7D !== 0 ? (diffRoom_7D / diffRN_7D) : 0);
    kpis.var7d.revpar = capacitatMes > 0 ? (diffRoom_7D / capacitatMes) : 0;

    var diffRN_7D_LY = kpis.ly.rn - acumulatsP7.rnLY;
    var diffRoom_7D_LY = kpis.ly.roomRev - acumulatsP7.roomRevLY;
    kpis.var7dLY.rn = diffRN_7D_LY;
    kpis.var7dLY.roomRev = diffRoom_7D_LY;
    kpis.var7dLY.fbRev = kpis.ly.fbRev - acumulatsP7.fbRevLY;
    kpis.var7dLY.totRev = kpis.ly.totRev - acumulatsP7.totRevLY;
    kpis.var7dLY.occ = capacitatMes > 0 ? (diffRN_7D_LY / capacitatMes) * 100 : 0;
    kpis.var7dLY.adr = (kpis.ly.rn > 0 && acumulatsP7.rnLY > 0) ? (kpis.ly.adr - (acumulatsP7.roomRevLY / acumulatsP7.rnLY)) : (diffRN_7D_LY !== 0 ? (diffRoom_7D_LY / diffRN_7D_LY) : 0);
    kpis.var7dLY.revpar = capacitatMes > 0 ? (diffRoom_7D_LY / capacitatMes) : 0;

    var tePressupost = false;
    if (shPress) {
      var dPress = shPress.getDataRange().getValues();
      var nomsMesosCurts = ["gener", "febrer", "març", "abril", "maig", "juny", "juliol", "agost", "setembre", "octubre", "novembre", "desembre"];
      var clauTextMes = nomsMesosCurts[mesIdx] + String(anySel).slice(-2);

      for (var p = 1; p < dPress.length; p++) {
        var rowP = dPress[p];
        if (normalitzarNomHotel(rowP[0]).toLowerCase() !== hotelSelNorm) continue;

        var mesVal = String(rowP[1] || "").trim().toLowerCase();
        if (mesVal === clauTextMes || dataYMD(rowP[1]).indexOf(mesClauSel) === 0) {
          if (String(rowP[2] || "").trim().toLowerCase() === "total") {
            kpis.pressupost.rn = Number(rowP[5]) || 0;
            kpis.pressupost.roomRev = Number(rowP[8]) || 0;
            kpis.pressupost.fbRev = Number(rowP[9]) || 0;
            kpis.pressupost.totRev = Number(rowP[10]) || (kpis.pressupost.roomRev + kpis.pressupost.fbRev);
            tePressupost = true;
            break;
          }
        }
      }

      if (tePressupost) {
        kpis.pressupost.occ = capacitatMes > 0 ? (kpis.pressupost.rn / capacitatMes) * 100 : 0;
        kpis.pressupost.adr = kpis.pressupost.rn > 0 ? (kpis.pressupost.roomRev / kpis.pressupost.rn) : 0;
        kpis.pressupost.revpar = capacitatMes > 0 ? (kpis.pressupost.roomRev / capacitatMes) : 0;

        kpis.varPress.rn = kpis.avui.rn - kpis.pressupost.rn;
        kpis.varPress.roomRev = kpis.avui.roomRev - kpis.pressupost.roomRev;
        kpis.varPress.fbRev = kpis.avui.fbRev - kpis.pressupost.fbRev;
        kpis.varPress.totRev = kpis.avui.totRev - kpis.pressupost.totRev;
        kpis.varPress.occ = kpis.avui.occ - kpis.pressupost.occ;
        kpis.varPress.adr = kpis.avui.adr - kpis.pressupost.adr;
        kpis.varPress.revpar = kpis.avui.revpar - kpis.pressupost.revpar;
      }
    }

    var teTancament = false;
    if (shTanc) {
      var dTanc = shTanc.getDataRange().getValues();
      var clauLYMes = (anySel >= 2026 ? (anySel - 1) : anySel) + "-" + partsM[1];

      for (var t = 1; t < dTanc.length; t++) {
        var rowT = dTanc[t];
        if (normalitzarNomHotel(rowT[0]).toLowerCase() !== hotelSelNorm) continue;
        if (String(rowT[3] || "").trim().toLowerCase() !== "total") continue;

        if (dataYMD(rowT[1]).indexOf(clauLYMes) === 0) {
          kpis.tancament.rn += Number(rowT[5]) || 0;
          kpis.tancament.roomRev += Number(rowT[7]) || 0;
          kpis.tancament.fbRev += Number(rowT[8]) || 0;
          kpis.tancament.totRev += Number(rowT[9]) || (Number(rowT[7]) + Number(rowT[8]));
          teTancament = true;
        }
      }
      if (teTancament) {
        kpis.tancament.occ = capacitatMes > 0 ? (kpis.tancament.rn / capacitatMes) * 100 : 0;
        kpis.tancament.adr = kpis.tancament.rn > 0 ? (kpis.tancament.roomRev / kpis.tancament.rn) : 0;
        kpis.tancament.revpar = capacitatMes > 0 ? (kpis.tancament.roomRev / capacitatMes) : 0;
      }
    }

    var canalsDef = [
      "Web Propia", "Particulars", "Agencies Online", "Agencies Online Net",
      "Altres Agencies", "Empreses Grups", "Empreses Agencia", "Col.lectius"
    ];

    function normCanal(txt) {
      return String(txt || "").toLowerCase().replace(/[\s_\.\-]/g, "");
    }

    var dadesCanals = {};
    canalsDef.forEach(function(cn) {
      dadesCanals[cn] = {
        otb: { rn: 0, rev: 0 },
        stly: { rn: 0, rev: 0 },
        p7: { rn: 0, rev: 0 },
        p7_stly: { rn: 0, rev: 0 }
      };
    });

    var totalRevFontsAct = 0;

    if (shFN) {
      var dFN = shFN.getDataRange().getValues();

      var mapaSnapsFN = {};
      for (var f0 = 1; f0 < dFN.length; f0++) {
        if (normalitzarNomHotel(dFN[f0][0]).toLowerCase() === hotelSelNorm) {
          var snIso = dataYMD(dFN[f0][1]);
          if (snIso) mapaSnapsFN[snIso] = true;
        }
      }
      var llistaSnapsFN = Object.keys(mapaSnapsFN).sort().reverse();
      
      var snapActFN = (llistaSnapsFN.indexOf(dataCapturaRef) !== -1) ? dataCapturaRef : (llistaSnapsFN.length > 0 ? llistaSnapsFN[0] : "");
      var snapP7_FN = trobarSnapEnLlista(snapActFN, 7, llistaSnapsFN);

      for (var f = 1; f < dFN.length; f++) {
        var rFN = dFN[f];
        if (normalitzarNomHotel(rFN[0]).toLowerCase() !== hotelSelNorm) continue;

        var stFN = dataYMD(rFN[2]);
        if (stFN.indexOf(mesClauSel) !== 0) continue;

        var snFN = dataYMD(rFN[1]);
        var nomCanalFila = normCanal(rFN[3]);

        var canalTrobat = "";
        for (var idxC = 0; idxC < canalsDef.length; idxC++) {
          if (normCanal(canalsDef[idxC]) === nomCanalFila) {
            canalTrobat = canalsDef[idxC];
            break;
          }
        }
        if (!canalTrobat) continue;

        var rnF = Number(rFN[4]) || 0;
        var revF = Number(rFN[5]) || 0;
        var rnFLY = Number(rFN[9]) || 0;
        var revFLY = Number(rFN[10]) || 0;

        if (snFN === snapActFN) {
          totalRevFontsAct += revF;
          dadesCanals[canalTrobat].otb.rn += rnF;
          dadesCanals[canalTrobat].otb.rev += revF;
          dadesCanals[canalTrobat].stly.rn += rnFLY;
          dadesCanals[canalTrobat].stly.rev += revFLY;
        } else if (snFN === snapP7_FN) {
          dadesCanals[canalTrobat].p7.rn += rnF;
          dadesCanals[canalTrobat].p7.rev += revF;
          dadesCanals[canalTrobat].p7_stly.rn += rnFLY;
          dadesCanals[canalTrobat].p7_stly.rev += revFLY;
        }
      }
    }

    var canalsRetorn = [];
    var webRevAct = 0;
    var webRevDiff7 = 0;

    var baseRoomRevAct = kpis.avui.roomRev > 0 ? kpis.avui.roomRev : totalRevFontsAct;
    var baseRoomRevAnt7D = kpis.avui.roomRev - kpis.var7d.roomRev;

    canalsDef.forEach(function(cn) {
      var objC = dadesCanals[cn];
      var pesOcup = kpis.avui.rn > 0 ? (objC.otb.rn / kpis.avui.rn) * 100 : 0;
      var pesProdc = baseRoomRevAct > 0 ? (objC.otb.rev / baseRoomRevAct) * 100 : 0;

      var adrOTB = objC.otb.rn > 0 ? (objC.otb.rev / objC.otb.rn) : 0;
      var adrSTLY = objC.stly.rn > 0 ? (objC.stly.rev / objC.stly.rn) : 0;

      var diffRN_PK = objC.otb.rn - objC.p7.rn;
      var diffRev_PK = objC.otb.rev - objC.p7.rev;
      var adrDiff_PK = diffRN_PK !== 0 ? (diffRev_PK / diffRN_PK) : 0;

      if (cn === "Web Propia") {
        webRevAct = objC.otb.rev;
        webRevDiff7 = diffRev_PK;
      }

      var diffRN_OTB_vs_STLY = objC.otb.rn - objC.stly.rn;
      var diffRev_OTB_vs_STLY = objC.otb.rev - objC.stly.rev;
      var adrDiff_OTB_vs_STLY = (objC.otb.rn > 0 && objC.stly.rn > 0) ? (adrOTB - adrSTLY) : (diffRN_OTB_vs_STLY !== 0 ? (diffRev_OTB_vs_STLY / diffRN_OTB_vs_STLY) : 0);

      var diffRN_PK_LY = objC.stly.rn - objC.p7_stly.rn;
      var diffRev_PK_LY = objC.stly.rev - objC.p7_stly.rev;
      var adrPK_LY = diffRN_PK_LY !== 0 ? (diffRev_PK_LY / diffRN_PK_LY) : 0;

      var varPKvsSTLY_RN = diffRN_PK - diffRN_PK_LY;
      var varPKvsSTLY_Rev = diffRev_PK - diffRev_PK_LY;
      var varPKvsSTLY_ADR = adrDiff_PK - adrPK_LY;

      canalsRetorn.push({
        canal: cn,
        pesOcup: pesOcup,
        pesProdc: pesProdc,
        otb: { rn: objC.otb.rn, rev: objC.otb.rev, adr: adrOTB },
        pk7: { rn: diffRN_PK, rev: diffRev_PK, adr: adrDiff_PK },
        varOTBvsSTLY: { rn: diffRN_OTB_vs_STLY, rev: diffRev_OTB_vs_STLY, adr: adrDiff_OTB_vs_STLY },
        varPKvsSTLY: { rn: varPKvsSTLY_RN, rev: varPKvsSTLY_Rev, adr: varPKvsSTLY_ADR }
      });
    });

    var webShareAct = baseRoomRevAct > 0 ? (webRevAct / baseRoomRevAct) * 100 : 0;
    var webRevFa7D = webRevAct - webRevDiff7;
    var webShareAntCalculat = baseRoomRevAnt7D > 0 ? (webRevFa7D / baseRoomRevAnt7D) * 100 : 0;

    var diffSenseFont = kpis.avui.roomRev - totalRevFontsAct;
    if (Math.abs(diffSenseFont) > 1) {
      canalsRetorn.push({
        canal: "Sense Font",
        esSenseFont: true,
        pesOcup: 0,
        pesProdc: baseRoomRevAct > 0 ? (diffSenseFont / baseRoomRevAct) * 100 : 0,
        otb: { rn: 0, rev: diffSenseFont, adr: 0 },
        pk7: { rn: 0, rev: 0, adr: 0 },
        varOTBvsSTLY: { rn: 0, rev: 0, adr: 0 },
        varPKvsSTLY: { rn: 0, rev: 0, adr: 0 }
      });
    }

    return {
      hotelActual: hotelSel,
      llistaHotels: llistaHotels,
      llistaSnaps: llistaSnaps,
      dataCapturaRef: dataCapturaRef,
      llistaMesos: llistaMesos,
      mesClauSel: mesClauSel,
      nomMesComplet: nomMesComplet,
      esMesEnCurs: esMesEnCurs,
      tePressupost: tePressupost,
      teTancament: teTancament,
      kpis: kpis,
      canals: canalsRetorn,
      satellits: {
        webShare: webShareAct,
        webShareAnt: webShareAntCalculat,
        senseFontRev: diffSenseFont
      }
    };
  } catch (err) {
    return { error: "Error Fitxa Mensual: " + err.message };
  }
}

// =========================================================================
// 5. GESTIÓ D'ESTRATÈGIA I NOTES (PERSISTÈNCIA PROTEGIDA A GOOGLE SHEETS)
// =========================================================================
function desarNotaEstrategia(hotel, dataEstadaISO, camp, valor) {
  if (!verificarPermisEdicioPricing()) {
    throw new Error("Accés denegat: només els editors autoritzats poden modificar dades de Pricing.");
  }
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sh = ss.getSheetByName("Estrategia_Notes");
    if (!sh) {
      sh = ss.insertSheet("Estrategia_Notes");
      sh.appendRow(["Hotel", "Data", "Esdeveniment", "Restriccions", "TarifaAnt", "TarifaAct", "ColorAct", "ColorAnt", "GrocAct", "GrocAnt", "EstatVenda"]);
    }
    var d = sh.getDataRange().getValues();
    var colMap = {
      esdeveniment: 3,
      restriccions: 4,
      tarifaAnt: 5,
      tarifaAct: 6,
      colorAct: 7,
      colorAnt: 8,
      grocAct: 9,
      grocAnt: 10,
      estatVenda: 11
    };
    var targetCol = colMap[camp];
    if (!targetCol) return;

    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim() === hotel && formatarDataISO(d[i][1]) === dataEstadaISO) {
        sh.getRange(i + 1, targetCol).setValue(valor);
        return;
      }
    }
    var newRow = [hotel, dataEstadaISO, "", "", "", "", "negre", "negre", false, false, "obert"];
    newRow[targetCol - 1] = valor;
    sh.appendRow(newRow);
  } catch (err) {
    console.error("Error desarNotaEstrategia:", err);
    throw err;
  }
}

function desarEstatVenda(hotel, dataEstadaISO, estat) {
  desarNotaEstrategia(hotel, dataEstadaISO, "estatVenda", estat);
}

function desarColorTarifa(hotel, dataEstadaISO, color) {
  desarNotaEstrategia(hotel, dataEstadaISO, "colorAct", color);
}

function desarFonsGrocTarifa(hotel, dataEstadaISO, estatGroc) {
  desarNotaEstrategia(hotel, dataEstadaISO, "grocAct", estatGroc);
}

function copiarTarifesActualsAAnteriors(hotelSel) {
  if (!verificarPermisEdicioPricing()) {
    throw new Error("Accés denegat: només els editors autoritzats poden copiar tarifes.");
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet() || (typeof SPREADSHEET_ID !== 'undefined' ? SpreadsheetApp.openById(SPREADSHEET_ID) : null);
    if (!ss) throw new Error("No s'ha pogut accedir al full de càlcul.");

    var shNotes = ss.getSheetByName("Estrategia_Notes");
    if (!shNotes) throw new Error("No s'ha trobat la pestanya 'Estrategia_Notes'.");

    var lastRow = shNotes.getLastRow();
    if (lastRow < 2) return { success: true, filesActualitzades: 0 };

    var dades = shNotes.getDataRange().getValues();
    var capcalera = dades[0];

    var colHotel = -1;
    var colTarifaAct = -1;
    var colTarifaAnt = -1;
    var colColorAct = -1;
    var colColorAnt = -1;
    var colGrocAct = -1;
    var colGrocAnt = -1;

    for (var c = 0; c < capcalera.length; c++) {
      var nomCol = String(capcalera[c] || "").trim().toLowerCase().replace(/[\s_\-\.]/g, "");
      if (nomCol === "hotel") colHotel = c;
      else if (nomCol === "tarifaactual") colTarifaAct = c;
      else if (nomCol === "tarifaanterior") colTarifaAnt = c;
      else if (nomCol === "coloractual") colColorAct = c;
      else if (nomCol === "coloranterior") colColorAnt = c;
      else if (nomCol === "fonsgrocactual") colGrocAct = c;
      else if (nomCol === "fonsgrocanterior") colGrocAnt = c;
    }

    if (colHotel === -1 || colTarifaAct === -1 || colTarifaAnt === -1) {
      throw new Error("No s'han identificat les columnes obligatòries a 'Estrategia_Notes'. Capçaleres detectades: " + capcalera.join(", "));
    }

    var hSelNorm = normalitzarNomHotel(hotelSel).toLowerCase();
    var filesActualitzades = 0;

    for (var r = 1; r < dades.length; r++) {
      var hFila = normalitzarNomHotel(dades[r][colHotel]).toLowerCase();
      if (hFila !== hSelNorm) continue;

      dades[r][colTarifaAnt] = dades[r][colTarifaAct];

      if (colColorAct !== -1 && colColorAnt !== -1) {
        dades[r][colColorAnt] = dades[r][colColorAct] || "negre";
        dades[r][colColorAct] = "negre";
      } else if (colColorAct !== -1) {
        dades[r][colColorAct] = "negre";
      }

      if (colGrocAct !== -1 && colGrocAnt !== -1) {
        dades[r][colGrocAnt] = dades[r][colGrocAct];
      }

      filesActualitzades++;
    }

    shNotes.getDataRange().setValues(dades);
    return { success: true, filesActualitzades: filesActualitzades };

  } catch (err) {
    throw new Error("Error en copiar tarifes matinals: " + err.message);
  }
}

// =========================================================================
// 6. INGESTA INTEGRADA DE BI2GO (FONTS DE NEGOCI I TIPOLOGIES)
// =========================================================================
function processarCorreuPrevisioDiaria() {
  var query = 'subject:"Informes diaris Bi2go" has:attachment newer_than:5d';
  var threads = GmailApp.search(query, 0, 5);
  if (threads.length === 0) return;

  var arxiuPrevisio = null;
  var arxiuTipus = null;

  for (var i = 0; i < threads.length; i++) {
    var missatges = threads[i].getMessages();
    for (var j = missatges.length - 1; j >= 0; j--) {
      var adjunts = missatges[j].getAttachments();
      for (var k = 0; k < adjunts.length; k++) {
        var nom = adjunts[k].getName().toLowerCase();
        if (nom.indexOf("previsión diaria tipo") !== -1 || nom.indexOf("prevision diaria tipo") !== -1) {
          arxiuTipus = adjunts[k];
        } else if (nom.indexOf("previsión diaria") !== -1 || nom.indexOf("prevision diaria") !== -1) {
          arxiuPrevisio = adjunts[k];
        }
      }
      if (arxiuPrevisio && arxiuTipus) break;
    }
    if (arxiuPrevisio && arxiuTipus) break;
  }

  var ssMestre = SpreadsheetApp.openById(SPREADSHEET_ID);

  if (arxiuPrevisio) {
    processarFitxerFontsNegoci(ssMestre, arxiuPrevisio);
  }

  if (arxiuTipus) {
    processarFitxerTipologies(ssMestre, arxiuTipus);
  }
}

function processarFitxerFontsNegoci(ssMestre, arxiuAdjunt) {
  var shTotal = assegurarPestanya(ssMestre, "Dades_Segments", "Segment");
  var shFonts = assegurarPestanya(ssMestre, "Dades_Fonts_Negoci", "Font_Negoci");

  var clausTotal = carregarClausExistents(shTotal);
  var clausFonts = carregarClausExistents(shFonts);

  var fitxerTemp = Drive.Files.insert(
    { title: "temp_bi2go_fn", mimeType: MimeType.GOOGLE_SHEETS },
    arxiuAdjunt.copyBlob(),
    { convert: true }
  );
  var tempSS = SpreadsheetApp.openById(fitxerTemp.id);
  var vals = tempSS.getSheets()[0].getDataRange().getValues();

  var novesTotal = [];
  var novesFonts = [];

  var fontsConfig = [
    { nom: "Web Propia",            colAct: 7,  colLY: 53 },
    { nom: "Particulars",           colAct: 11, colLY: 57 },
    { nom: "Agencies Online",       colAct: 15, colLY: 61 },
    { nom: "Agencies Online Net",   colAct: 19, colLY: 65 },
    { nom: "Altres Agencies",       colAct: 23, colLY: 69 },
    { nom: "Empreses Grups",        colAct: 27, colLY: 73 },
    { nom: "Empreses Agencia",      colAct: 31, colLY: 77 },
    { nom: "Col.lectius",           colAct: 35, colLY: 81 }
  ];

  for (var fIdx = 2; fIdx < vals.length; fIdx++) {
    var row = vals[fIdx];
    var nomHotel = normalitzarNomHotel(row[0]);
    if (!nomHotel) continue;

    var snapshotDate = formatarDataISO(row[1]);
    var stayDate = formatarDataISO(row[2]);
    if (!snapshotDate || !stayDate) continue;

    var clauTotal = nomHotel + "_" + snapshotDate + "_" + stayDate + "_Total";
    if (!clausTotal[clauTotal]) {
      var rn = Number(row[3]) || 0;
      var roomRev = Number(row[4]) || 0;
      var fbRev = Number(row[5]) || 0;
      var totRev = Number(row[6]) || (roomRev + fbRev);
      var adr = rn > 0 ? (roomRev / rn) : 0;

      var rnLY = Number(row[49]) || 0;
      var roomRevLY = Number(row[50]) || 0;
      var fbRevLY = Number(row[51]) || 0;
      var totRevLY = Number(row[52]) || (roomRevLY + fbRevLY);
      var adrLY = rnLY > 0 ? (roomRevLY / rnLY) : 0;

      novesTotal.push([
        nomHotel, snapshotDate, stayDate, "Total",
        rn, roomRev, fbRev, totRev, adr,
        rnLY, roomRevLY, fbRevLY, totRevLY, adrLY
      ]);
      clausTotal[clauTotal] = true;
    }

    fontsConfig.forEach(function(fn) {
      var clauFN = nomHotel + "_" + snapshotDate + "_" + stayDate + "_" + fn.nom;
      if (!clausFonts[clauFN]) {
        var cA = fn.colAct;
        var fRn = Number(row[cA]) || 0;
        var fRoom = Number(row[cA + 1]) || 0;
        var fFb = Number(row[cA + 2]) || 0;
        var fTot = Number(row[cA + 3]) || (fRoom + fFb);
        var fAdr = fRn > 0 ? (fRoom / fRn) : 0;

        var cL = fn.colLY;
        var fRnLY = Number(row[cL]) || 0;
        var fRoomLY = Number(row[cL + 1]) || 0;
        var fFbLY = Number(row[cL + 2]) || 0;
        var fTotLY = Number(row[cL + 3]) || (fRoomLY + fFbLY);
        var fAdrLY = fRnLY > 0 ? (fRoomLY / fRnLY) : 0;

        novesFonts.push([
          nomHotel, snapshotDate, stayDate, fn.nom,
          fRn, fRoom, fFb, fTot, fAdr,
          fRnLY, fRoomLY, fFbLY, fTotLY, fAdrLY
        ]);
        clausFonts[clauFN] = true;
      }
    });
  }

  Drive.Files.remove(fitxerTemp.id);

  if (novesTotal.length > 0) {
    shTotal.getRange(shTotal.getLastRow() + 1, 1, novesTotal.length, novesTotal[0].length).setValues(novesTotal);
  }
  if (novesFonts.length > 0) {
    shFonts.getRange(shFonts.getLastRow() + 1, 1, novesFonts.length, novesFonts[0].length).setValues(novesFonts);
  }
}

function processarFitxerTipologies(ssMestre, arxiuAdjunt) {
  var shTipus = assegurarPestanya(ssMestre, "Dades_Tipologies", "Tipologia");
  var clausTipus = carregarClausExistents(shTipus);

  var fitxerTemp = Drive.Files.insert(
    { title: "temp_bi2go_tipus", mimeType: MimeType.GOOGLE_SHEETS },
    arxiuAdjunt.copyBlob(),
    { convert: true }
  );
  var tempSS = SpreadsheetApp.openById(fitxerTemp.id);
  var vals = tempSS.getSheets()[0].getDataRange().getValues();
  var novesTipus = [];

  var tipusConfig = [
    { nom: "Confort",               colAct: 7,  colLY: 59 },
    { nom: "Confort Vistes/Jardí",   colAct: 11, colLY: 63 },
    { nom: "Familiar",              colAct: 15, colLY: 67 },
    { nom: "Familiar Vistes/Jardí", colAct: 19, colLY: 71 },
    { nom: "Confort especial",      colAct: 23, colLY: 75 },
    { nom: "Familiar especial",     colAct: 27, colLY: 79 },
    { nom: "Superior",              colAct: 31, colLY: 83 },
    { nom: "Suite",                 colAct: 35, colLY: 87 },
    { nom: "Suite Exclusiva",       colAct: 39, colLY: 91 },
    { nom: "Economica",             colAct: 43, colLY: 95 }
  ];

  for (var fIdx = 2; fIdx < vals.length; fIdx++) {
    var row = vals[fIdx];
    var nomHotel = normalitzarNomHotel(row[0]);
    if (!nomHotel) continue;

    var snapshotDate = formatarDataISO(row[1]);
    var stayDate = formatarDataISO(row[2]);
    if (!snapshotDate || !stayDate) continue;

    tipusConfig.forEach(function(tc) {
      var clauT = nomHotel + "_" + snapshotDate + "_" + stayDate + "_" + tc.nom;
      if (!clausTipus[clauT]) {
        var cA = tc.colAct;
        var tRn = Number(row[cA]) || 0;
        var tRoom = Number(row[cA + 1]) || 0;
        var tFb = Number(row[cA + 2]) || 0;
        var tTot = Number(row[cA + 3]) || (tRoom + tFb);
        var tAdr = tRn > 0 ? (tRoom / tRn) : 0;

        var cL = tc.colLY;
        var tRnLY = Number(row[cL]) || 0;
        var tRoomLY = Number(row[cL + 1]) || 0;
        var tFbLY = Number(row[cL + 2]) || 0;
        var tTotLY = Number(row[cL + 3]) || (tRoomLY + tFbLY);
        var tAdrLY = tRnLY > 0 ? (tRoomLY / tRnLY) : 0;

        novesTipus.push([
          nomHotel, snapshotDate, stayDate, tc.nom,
          tRn, tRoom, tFb, tTot, tAdr,
          tRnLY, tRoomLY, tFbLY, tTotLY, tAdrLY
        ]);
        clausTipus[clauT] = true;
      }
    });
  }

  Drive.Files.remove(fitxerTemp.id);

  if (novesTipus.length > 0) {
    shTipus.getRange(shTipus.getLastRow() + 1, 1, novesTipus.length, novesTipus[0].length).setValues(novesTipus);
  }
}

function assegurarPestanya(ss, nomPestanya) {
  var sheet = ss.getSheetByName(nomPestanya);
  if (!sheet) {
    sheet = ss.insertSheet(nomPestanya);
  }
  return sheet;
}

function carregarClausExistents(sheet, colClau1, colClau2, colClau3) {
  var claus = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return claus;
  
  var dades = sheet.getDataRange().getValues();
  for (var i = 1; i < dades.length; i++) {
    var c1 = colClau1 !== undefined ? String(dades[i][colClau1] || "").trim() : "";
    var c2 = colClau2 !== undefined ? String(dades[i][colClau2] || "").trim() : "";
    var c3 = colClau3 !== undefined ? String(dades[i][colClau3] || "").trim() : "";
    var clau = c1 + "_" + c2 + (c3 ? "_" + c3 : "");
    if (clau !== "_") {
      claus[clau] = true;
    }
  }
  return claus;
}
function ferBackupAutomaticSheet() {
  var idOriginal = SPREADSHEET_ID;
  var nomCarpetaBackup = "Backups_Revenue_Manager";
  var maxCopies = 8; // Conserva les últimes 8 còpies (ex. 2 mesos si és setmanal)

  var fitxer = DriveApp.getFileById(idOriginal);
  var carpetes = DriveApp.getFoldersByName(nomCarpetaBackup);
  var carpeta = carpetes.hasNext() ? carpetes.next() : DriveApp.createFolder(nomCarpetaBackup);

  var dataAvui = Utilities.formatDate(new Date(), "Europe/Madrid", "yyyy-MM-dd_HHmm");
  fitxer.makeCopy("BACKUP_Revenue_" + dataAvui, carpeta);

  // Neteja de còpies antigues per no omplir Drive
  var arxius = [];
  var iter = carpeta.getFiles();
  while (iter.hasNext()) {
    arxius.push(iter.next());
  }
  arxius.sort(function(a, b) {
    return b.getDateCreated().getTime() - a.getDateCreated().getTime();
  });
  for (var i = maxCopies; i < arxius.length; i++) {
    arxius[i].setTrashed(true);
  }
}
