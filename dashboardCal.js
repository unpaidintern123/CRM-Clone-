function getCampaignName(){
  const sheet = SpreadsheetApp.openById("1fiqGa4nRkSkun-nffI_QVtFQG4o4ubJGXkMIcEprI1k").getSheetByName("Leads Data");

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.splice(1);
  Logger.log(rows);
  const campaignNameIndex = headers.indexOf("CampaignName");
  const adNameIndex = headers.indexOf("AdName");
  const campaignSourceIndex = headers.indexOf("CampaignSource");
  const placementIndex = headers.indexOf("Placement");
  const adSetName = headers.indexOf("AdSetName");

  const campaignNames = [...new Set(rows.map((row) => row[campaignNameIndex]).filter((val) => val !== ""))];
  const adNames = [...new Set(rows.map((row) => row[adNameIndex]).filter((val)=> val !== ""))]
  const platforms = [...new Set(rows.map((row) => row[campaignSourceIndex]).filter((val)=> val !== ""))]
  const placements = [...new Set(rows.map((row) => row[placementIndex]).filter((val)=> val !== ""))]
  const adSetNames = [...new Set(rows.map((row) => row[adSetName]).filter((val)=> val !== ""))]

  return {
    campaignNames: campaignNames,
    adNames: adNames,
    platforms: platforms,
    placements: placements,
    adSetNames : adSetNames
  }
}

function retriveLeadData2(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Leads Data");
  const data = sheet.getDataRange().getValues();
  const headings = data[0]
  const leadData = data.splice(1);
  const leadDataBase = []
  leadData.forEach((lead)=>{
    let obj = {}
    headings.forEach((heading,index)=>{
      obj[heading.toLowerCase()] = lead[index];
    })
    leadDataBase.push(obj)
  })
  return JSON.stringify(leadDataBase);
  }

function retriveConsData(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Consultation Booked Logs");
  const data = sheet.getDataRange().getValues();
  const headings = data[0]
  const leadData = data.splice(1);
  const leadDataBase = []
  leadData.forEach((lead)=>{
    let obj = {}
    headings.forEach((heading,index)=>{
      obj[heading.toLowerCase()] = lead[index];
    })
    leadDataBase.push(obj)
  })
  return JSON.stringify(leadDataBase);
  }


function getLatestWeekEndData() {
  const sheet = SpreadsheetApp.openById("1seVqJ5xPqva1Si6CXDpuZSePvP40PDO2fzJSJfss8GU").getSheetByName("MIS Data");
  const allData = sheet.getDataRange().getValues();
  const headers = allData[0];
  const rows = allData.slice(1);

  const dateColIndex = headers.indexOf("Week End Date");
  if (dateColIndex === -1) {
    throw new Error('"Week End Date" column not found.');
  }

  const today = new Date();
  const parsedData = [];

  // Collect rows with valid date and calculate date difference
  rows.forEach(row => {
    let dateCell = row[dateColIndex];
    let parsedDate;

    if (dateCell instanceof Date) {
      parsedDate = dateCell;
    } else if (typeof dateCell === 'string' && dateCell.includes('/')) {
      const parts = dateCell.split('/');
      if (parts.length === 3) {
        parsedDate = new Date(`${parts[2]}-${parts[1]}-${parts[0]}`); // yyyy-MM-dd
      }
    }

    if (parsedDate && !isNaN(parsedDate)) {
      const diff = Math.abs(parsedDate - today);
      parsedData.push({ row, parsedDate, diff });
    }
  });

  if (parsedData.length === 0) return [];

  // Find the minimum difference
  parsedData.sort((a, b) => a.diff - b.diff);
  const closestDate = parsedData[0].parsedDate;

  // Filter rows with that closest date
  const result = parsedData
    .filter(entry => entry.parsedDate.getTime() === closestDate.getTime())
    .map(entry => entry.row);

  return JSON.stringify(result);
}


