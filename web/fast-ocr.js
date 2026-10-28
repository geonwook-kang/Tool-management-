/* Pixel-only label proposals. No names, serials or expected dates are inputs. */
(function (global) {
  'use strict';
  const median = (values) => { const sorted = values.slice().sort((a,b) => a-b); return sorted.length ? sorted[Math.floor(sorted.length/2)] : 0; };
  const area = (box) => Math.max(0, box.x1-box.x0)*Math.max(0, box.y1-box.y0);
  const overlap = (a,b) => Math.max(0,Math.min(a.x1,b.x1)-Math.max(a.x0,b.x0))*Math.max(0,Math.min(a.y1,b.y1)-Math.max(a.y0,b.y0));
  function components(pixels, width, height, threshold) {
    const mask = new Uint8Array(width*height), pending = new Int32Array(width*height), result=[];
    for(let at=0;at<mask.length;at++) mask[at]=pixels[at*4+3]>128&&(pixels[at*4]+pixels[at*4+1]+pixels[at*4+2])/3>=threshold?1:0;
    for(let start=0;start<mask.length;start++) {
      if(!mask[start]) continue;
      let size=1,next=0,x0=width,x1=0,y0=height,y1=0;pending[0]=start;mask[start]=0;
      while(next<size) {
        const at=pending[next++],x=at%width,y=Math.floor(at/width);x0=Math.min(x0,x);x1=Math.max(x1,x);y0=Math.min(y0,y);y1=Math.max(y1,y);
        for(const n of [x>0?at-1:-1,x<width-1?at+1:-1,y>0?at-width:-1,y<height-1?at+width:-1]) if(n>=0&&mask[n]) {mask[n]=0;pending[size++]=n;}
      }
      const w=x1-x0+1,h=y1-y0+1,box={x0,y0,x1:x1+1,y1:y1+1};
      if(w<width*.055||h<height*.014||w*h<width*height*.0014||w*h>width*height*.28||w/h<.65||w/h>10||size/(w*h)<.4) continue;
      const edge=x0===0||y0===0||x1===width-1||y1===height-1;
      // A metal stem may connect a real serial tag to the top of the photo.
      // A narrow top component is still useful; a bright room/floor is not.
      if(edge&&!(y0===0&&x0>0&&x1<width-1&&w<width*.55&&h<height*.5)) continue;
      result.push({bbox:box,width:w,height:h,fill:size/(w*h),threshold,edge});
    }
    return result;
  }
  function regions(canvas) {
    const source=canvas.nativeCanvas||canvas, scale=Math.min(1,850/Math.max(source.width,source.height));
    const sample=document.createElement('canvas');sample.width=Math.max(1,Math.round(source.width*scale));sample.height=Math.max(1,Math.round(source.height*scale));
    const context=sample.getContext('2d',{willReadFrequently:true});context.drawImage(source,0,0,sample.width,sample.height);
    const width=sample.width,height=sample.height,pixels=context.getImageData(0,0,width,height).data;
    const candidates=[115,135,160,180,205,225].flatMap(threshold=>components(pixels,width,height,threshold));
    const columns=[];
    for(const threshold of [115,135,160,180]) {
      const cells=candidates.filter(c=>c.threshold===threshold&&c.width/c.height>2.3&&c.width/c.height<8&&c.height<height*.12&&c.width>width*.15&&!c.edge);
      for(const first of cells) {
        const group=cells.filter(c=>Math.abs(c.bbox.x0-first.bbox.x0)<Math.max(c.height,first.height)*.55&&Math.abs(c.width-first.width)<Math.max(c.width,first.width)*.35).sort((a,b)=>a.bbox.y0-b.bbox.y0);
        const run=[];
        for(const cell of group) {
          const previous=run[run.length-1];
          if(previous&&(cell.bbox.y0+cell.bbox.y1-previous.bbox.y0-previous.bbox.y1)/2>Math.max(cell.height,previous.height)*1.7) { if(run.length>=2) break;run.length=0; }
          run.push(cell);
        }
        if(run.length<2||run.length>5) continue;
        const bbox={x0:Math.min(...run.map(c=>c.bbox.x0)),y0:Math.min(...run.map(c=>c.bbox.y0)),x1:Math.max(...run.map(c=>c.bbox.x1)),y1:Math.max(...run.map(c=>c.bbox.y1))};
        if(!columns.some(c=>overlap(c.bbox,bbox)/Math.min(area(c.bbox),area(bbox))>.75)) columns.push({kind:'dates',bbox,cells:run.map(c=>c.bbox),score:run.length*2+Math.sqrt(area(bbox)/(width*height)),threshold});
      }
    }
    columns.sort((a,b)=>b.score-a.score||b.bbox.x0-a.bbox.x0);
    const tags=candidates.filter(c=>c.height>height*.065&&c.width>width*.1&&c.width/c.height<5&&!(c.width/c.height>3&&c.height<height*.1)).map(c=>({kind:'tag',bbox:c.bbox,score:Math.sqrt(area(c.bbox)/(width*height))*4+c.fill+(c.width/c.height>=1.15?.4:0)+(c.threshold>=205?.2:0),threshold:c.threshold})).sort((a,b)=>b.score-a.score);
    const selected=[];
    for(const tag of tags) {
      if(columns.slice(0,1).some(c=>overlap(c.bbox,tag.bbox)/Math.min(area(c.bbox),area(tag.bbox))>.8)) continue;
      if(selected.some(c=>overlap(c.bbox,tag.bbox)/Math.min(area(c.bbox),area(tag.bbox))>.65)) continue;
      selected.push(tag);if(selected.length===3) break;
    }
    const table=columns.slice(0,1);
    for(const column of table) {
      // A dark table rule may join one white cell to its neighbour. Retain
      // source pixels one measured row below a two-cell proposal as well;
      // this is only a crop extension, never an inferred date or role.
      if(column.cells.length===2) {
        const first=column.cells[0],last=column.cells[1],step=(last.y0+last.y1-first.y0-first.y1)/2;
        const next={x0:last.x0,x1:last.x1,y0:last.y0+step,y1:Math.min(height,last.y1+step)};
        if(step>0&&next.y1>next.y0) {column.cells.push(next);column.bbox.y1=next.y1;column.extended=true;}
      }
    }
    const result=selected.slice(0,2).sort((a,b)=>a.bbox.y0-b.bbox.y0).concat(table);
    const map=box=>({x0:box.x0/scale,y0:box.y0/scale,x1:box.x1/scale,y1:box.y1/scale});
    return result.map(region=>Object.assign({},region,{bbox:map(region.bbox),cells:region.cells&&region.cells.map(map),sourceWidth:source.width,sourceHeight:source.height}));
  }
  function otsu(image) {
    const pixels=image.data,histogram=new Uint32Array(256);let total=0,sum=0;
    for(let at=0;at<pixels.length;at+=4) {const level=Math.round((pixels[at]+pixels[at+1]+pixels[at+2])/3);histogram[level]++;total++;sum+=level;}
    let count=0,left=0,best=-1,threshold=0;
    for(let level=0;level<255;level++) {count+=histogram[level];left+=level*histogram[level];if(!count||count===total) continue;const variance=count*(total-count)*Math.pow(left/count-(sum-left)/(total-count),2);if(variance>best) {best=variance;threshold=level;}}
    return best>0?threshold:null;
  }
  function trimRules(canvas) {
    const width=canvas.width,height=canvas.height,pixels=canvas.getContext('2d',{willReadFrequently:true}).getImageData(0,0,width,height).data;
    const mask=new Uint8Array(width*height),pending=new Int32Array(width*height);let top=0,bottom=height;
    for(let at=0;at<mask.length;at++) mask[at]=(pixels[at*4]+pixels[at*4+1]+pixels[at*4+2])/3<140?1:0;
    for(let start=0;start<mask.length;start++) {
      if(!mask[start]) continue;let size=1,next=0,x0=width,x1=0,y0=height,y1=0;pending[0]=start;mask[start]=0;
      while(next<size) {const at=pending[next++],x=at%width,y=Math.floor(at/width);x0=Math.min(x0,x);x1=Math.max(x1,x);y0=Math.min(y0,y);y1=Math.max(y1,y);
        for(const n of [x>0?at-1:-1,x<width-1?at+1:-1,y>0?at-width:-1,y<height-1?at+width:-1]) if(n>=0&&mask[n]) {mask[n]=0;pending[size++]=n;}}
      if(x1-x0<width*.15||(x1-x0)/Math.max(1,y1-y0)<5||y1-y0>height*.3||(y1-y0>height*.12&&size/Math.max(1,(x1-x0)*(y1-y0))>.65)) continue;
      if(y1<height*.3) top=Math.max(top,y1+2);
      if(y0>height*.7) bottom=Math.min(bottom,y0-2);
    }
    if(bottom-top<height*.55||(!top&&bottom===height)) return canvas;
    const output=document.createElement('canvas');output.width=width;output.height=bottom-top;output.getContext('2d').drawImage(canvas,0,top,width,bottom-top,0,0,width,bottom-top);return output;
  }
  function straightenLine(canvas) {
    // Estimate a baseline from connected ink glyphs, never their meanings.
    // Small perspective slopes otherwise lower confidence even in clear rows.
    const width=canvas.width,height=canvas.height,pixels=canvas.getContext('2d',{willReadFrequently:true}).getImageData(0,0,width,height).data;
    const mask=new Uint8Array(width*height),pending=new Int32Array(width*height),glyphs=[];
    for(let at=0;at<mask.length;at++) mask[at]=(pixels[at*4]+pixels[at*4+1]+pixels[at*4+2])/3<140?1:0;
    for(let start=0;start<mask.length;start++) {
      if(!mask[start]) continue;let size=1,next=0,x0=width,x1=0,y0=height,y1=0;pending[0]=start;mask[start]=0;
      while(next<size) {const at=pending[next++],x=at%width,y=Math.floor(at/width);x0=Math.min(x0,x);x1=Math.max(x1,x);y0=Math.min(y0,y);y1=Math.max(y1,y);
        for(const n of [x>0?at-1:-1,x<width-1?at+1:-1,y>0?at-width:-1,y<height-1?at+width:-1]) if(n>=0&&mask[n]) {mask[n]=0;pending[size++]=n;}}
      if(y1-y0>height*.3&&y1-y0<height*.94&&x1-x0>2&&x1-x0<width*.2) glyphs.push({x:(x0+x1)/2,y:y1,h:y1-y0});
    }
    if(glyphs.length<5||Math.max(...glyphs.map(g=>g.x))-Math.min(...glyphs.map(g=>g.x))<width*.45) return canvas;
    const fit=points=>{const mx=points.reduce((n,p)=>n+p.x,0)/points.length,my=points.reduce((n,p)=>n+p.y,0)/points.length;let numerator=0,denominator=0;for(const p of points){numerator+=(p.x-mx)*(p.y-my);denominator+=(p.x-mx)*(p.x-mx);}return {mx,my,slope:denominator?numerator/denominator:0};};
    let line=fit(glyphs),h=median(glyphs.map(g=>g.h));
    const clean=glyphs.filter(g=>Math.abs(g.y-(line.my+line.slope*(g.x-line.mx)))<h*.22);
    if(clean.length>=5) line=fit(clean);
    if(!Number.isFinite(line.slope)||Math.abs(line.slope)>.35||Math.abs(line.slope)*width<3) return canvas;
    const padding=Math.ceil(Math.abs(line.slope)*width/2)+2,result=document.createElement('canvas');result.width=width;result.height=height+padding*2;
    const context=result.getContext('2d');context.fillStyle='#fff';context.fillRect(0,0,result.width,result.height);
    for(let x=0;x<width;x++) context.drawImage(canvas,x,0,1,height,x,padding-Math.round(line.slope*(x-width/2)),1,height);
    return result;
  }
  function prepareWholeColumn(canvas, region) {
    const source=canvas.nativeCanvas||canvas,box=region.bbox,w=box.x1-box.x0;
    const rowHeight=median((region.cells||[box]).map(cell=>cell.y1-cell.y0));
    const x=Math.max(0,Math.floor(box.x0+w*.025)),y=Math.max(0,Math.floor(box.y0-rowHeight*.1));
    const width=Math.max(1,Math.min(source.width,Math.ceil(box.x1-w*.025))-x),height=Math.max(1,Math.min(source.height,Math.ceil(box.y1+rowHeight*.1))-y);
    const output=document.createElement('canvas');output.width=width;output.height=height;output.getContext('2d').drawImage(source,x,y,width,height,0,0,width,height);return output;
  }
  function prepare(canvas, region, options) {
    const source=canvas.nativeCanvas||canvas, settings=options||{}, boxes=region.cells||[region.bbox], strips=[];
    for(const box of boxes) {
      const w=box.x1-box.x0,h=box.y1-box.y0;
      const inset=region.kind==='dates'?Math.max(1,h*.07):0;
      const x=Math.max(0,Math.floor(box.x0+inset)),y=Math.max(0,Math.floor(box.y0+inset));
      const width=Math.max(1,Math.min(source.width,Math.ceil(box.x1-inset))-x),height=Math.max(1,Math.min(source.height,Math.ceil(box.y1-inset))-y);
      let strip=document.createElement('canvas');strip.width=width;strip.height=height;
      strip.getContext('2d').drawImage(source,x,y,width,height,0,0,width,height);
      if(region.kind==='dates') strip=straightenLine(trimRules(strip));
      const ctx=strip.getContext('2d',{willReadFrequently:true});
      if(settings.binary) {const im=ctx.getImageData(0,0,strip.width,strip.height),threshold=otsu(im);if(threshold!=null) {for(let at=0;at<im.data.length;at+=4) {const value=(im.data[at]+im.data[at+1]+im.data[at+2])/3<=threshold?0:255;im.data[at]=im.data[at+1]=im.data[at+2]=value;}ctx.putImageData(im,0,0);}}
      strips.push(strip);
    }
    const padding=20,gap=30,output=document.createElement('canvas');output.width=Math.max(...strips.map(s=>s.width))+padding*2;output.height=strips.reduce((n,s)=>n+s.height,0)+padding*2+gap*(strips.length-1);
    const context=output.getContext('2d');context.fillStyle='#fff';context.fillRect(0,0,output.width,output.height);let y=padding;
    for(const strip of strips) {context.drawImage(strip,padding,y);y+=strip.height+gap;}
    return output;
  }
  global.ToolFastOCR=Object.freeze({regions,prepare,prepareWholeColumn});
})(typeof window!=='undefined'?window:globalThis);
