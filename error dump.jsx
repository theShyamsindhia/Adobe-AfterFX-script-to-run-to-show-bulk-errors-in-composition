/* IMAI_Expression_Diagnostics.jsx
   Scans expressions across comps, layers, and properties.
   - Forces evaluation, collects expressionError, property paths, and expression snippets.
   - Detects missing thisComp.layer("...") references.
   - Warns on probable 2D/3D scale mismatches.
   - Outputs a copyable dialog + saves a .txt on Desktop.

   Usage: File > Scripts > Run Script File...
*/

(function () {
  app.beginUndoGroup("IMAI – Expression Diagnostics");

  // ------------ UI: scope ------------
  var scopeAll = false;
  (function askScope(){
    var w = new Window("dialog","Scan Scope");
    w.add("statictext",undefined,"Scan expressions for:");
    var rb1 = w.add("radiobutton",undefined,"Active Comp only");
    var rb2 = w.add("radiobutton",undefined,"All Comps in project");
    rb1.value = true;
    var g = w.add("group");
    var ok = g.add("button",undefined,"Run",{name:"ok"});
    var cancel = g.add("button",undefined,"Cancel");
    ok.onClick = function(){ scopeAll = rb2.value; w.close(1); };
    cancel.onClick = function(){ w.close(0); };
    if (w.show() !== 1) { throw "User cancelled."; }
  })();

  // ------------ helpers ------------
  function pad(n){ return (n<10?"0":"")+n; }
  function timestamp(){
    var d=new Date();
    return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate())+"_"+pad(d.getHours())+"-"+pad(d.getMinutes())+"-"+pad(d.getSeconds());
  }
  function esc(s){ s = String(s||""); return s.replace(/\r/g,"\\r").replace(/\n/g,"\\n"); }

  function getComps(){
    var arr=[];
    if (!scopeAll) {
      var ac = app.project.activeItem;
      if (ac && ac instanceof CompItem) arr.push(ac);
    } else {
      for (var i=1;i<=app.project.numItems;i++){
        var it = app.project.item(i);
        if (it instanceof CompItem) arr.push(it);
      }
    }
    return arr;
  }

  function propPath(prop){
    var names=[];
    var p=prop;
    while (p && p.parentProperty){
      names.unshift(p.name);
      p = p.parentProperty;
    }
    var lyr = prop.propertyGroup(prop.propertyDepth);
    return names.join(" > ");
  }

  // simple regex to find thisComp.layer("NAME") and thisComp.layer('NAME')
  var reLayerCall = /thisComp\s*\.\s*layer\s*\(\s*["']([^"']+)["']\s*\)/g;

  function listMissingLayerRefs(exprText, comp){
    var miss=[], m;
    reLayerCall.lastIndex = 0;
    while ((m = reLayerCall.exec(exprText)) !== null){
      var name = m[1];
      var found = false;
      for (var i=1;i<=comp.numLayers;i++){
        if (comp.layer(i).name === name) { found=true; break; }
      }
      if (!found && miss.indexOf(name)===-1) miss.push(name);
    }
    return miss;
  }

  function firstLines(s, maxLines){
    var parts = String(s||"").split(/\r?\n/);
    var keep = parts.slice(0, maxLines);
    var txt = keep.join("\\n");
    if (parts.length > maxLines) txt += " ...";
    return txt;
  }

  function tryEvaluate(prop, t){
    // Touch value to force evaluation; guard for properties that throw
    try {
      if (prop.propertyValueType !== PropertyValueType.NO_VALUE) {
        // For some props, valueAtTime is safer than value
        prop.valueAtTime(t, false);
      }
    } catch (e) {
      // We don't rethrow; expressionError should capture it
    }
  }

  function scanPropertyGroup(comp, layer, group, out, t){
    for (var i=1; i<=group.numProperties; i++){
      var p = group.property(i);
      if (!p) continue;

      if (p.propertyType === PropertyType.PROPERTY) {
        var hasExpr = false, exprText = "";
        try {
          hasExpr = p.canSetExpression && (p.expression !== "" || p.expressionEnabled);
          exprText = p.expression || "";
        } catch(e) {}

        if (hasExpr) {
          // Force evaluation to populate expressionError
          tryEvaluate(p, t);

          var err = "";
          try { err = p.expressionError || ""; } catch(e){ err=""; }

          // Detect missing layer refs
          var missingRefs = listMissingLayerRefs(exprText, comp);

          // Heuristic: Scale 3D vs 2D mismatch warning if there is an error
          var scaleWarn = "";
          try {
            if (p.matchName === "ADBE Scale" && err) {
              // If the layer is 3D, AE expects [x,y,z]
              if (layer.threeDLayer) {
                scaleWarn = "Layer is 3D: Scale expects [x,y,z]. Expression may be returning 2D.";
              } else {
                // If layer 2D but plugin/effect forces 3D array; rarer
                scaleWarn = "";
              }
            }
          } catch(e){}

          if (err || missingRefs.length || scaleWarn) {
            out.push({
              comp: comp.name,
              layerIndex: layer.index,
              layerName: layer.name,
              propPath: propPath(p),
              matchName: p.matchName,
              error: err,
              missing: missingRefs,
              scaleWarn: scaleWarn,
              exprFirst: firstLines(exprText, 8)
            });
          }
        }
      } else if (p.propertyType === PropertyType.INDEXED_GROUP || p.propertyType === PropertyType.NAMED_GROUP) {
        scanPropertyGroup(comp, layer, p, out, t);
      }
    }
  }

  // ------------ Scan ------------
  var results = [];
  var comps = getComps();
  if (comps.length === 0) { alert("No comp to scan. Open/select a comp and try again."); return; }

  for (var c=0;c<comps.length;c++){
    var comp = comps[c];
    // Evaluate at mid-time of comp to catch time-dependent errors
    var t = Math.min(Math.max(0.0, comp.duration*0.5), Math.max(0.0, comp.duration - 0.001));
    for (var li=1; li<=comp.numLayers; li++){
      var L = comp.layer(li);
      scanPropertyGroup(comp, L, L, results, t);
    }
  }

  // ------------ Report ------------
  var lines = [];
  lines.push("After Effects Expression Diagnostics");
  lines.push("App Version: " + app.version);
  lines.push("Timestamp: " + timestamp());
  lines.push("Scope: " + (scopeAll ? "All Comps" : "Active Comp"));
  lines.push("----------------------------------------");

  if (results.length === 0) {
    lines.push("No expression issues found. ✅");
  } else {
    for (var i=0;i<results.length;i++){
      var r = results[i];
      lines.push("#"+(i+1));
      lines.push("Comp:      " + r.comp);
      lines.push("Layer ["+r.layerIndex+"]: " + r.layerName);
      lines.push("Property:  " + r.propPath + "  ("+r.matchName+")");
      if (r.error)        lines.push("Error:     " + r.error);
      if (r.missing && r.missing.length) lines.push("Missing thisComp.layer refs: " + r.missing.join(", "));
      if (r.scaleWarn)    lines.push("Note:      " + r.scaleWarn);
      if (r.exprFirst)    lines.push("Expr Snip: " + r.exprFirst);
      lines.push("----------------------------------------");
    }
  }

  var report = lines.join("\n");

  // Save to Desktop
  var outPath = Folder.desktop.fsName + "/AE_Expression_Diagnostics_" + timestamp() + ".txt";
  try {
    var f = new File(outPath);
    f.encoding = "UTF-8";
    if (f.open("w")) { f.write(report); f.close(); }
  } catch(e) {}

  // Show copyable dialog
  var dlg = new Window("dialog", "Expression Diagnostics");
  var edit = dlg.add("edittext", undefined, report, {multiline:true, scrolling:true});
  edit.preferredSize = [900, 600];
  var g = dlg.add("group");
  g.alignment = "right";
  g.add("statictext", undefined, "Saved: " + outPath);
  var ok = g.add("button", undefined, "Close", {name:"ok"});
  dlg.show();

  app.endUndoGroup();
})();
