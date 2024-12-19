document.addEventListener("DOMContentLoaded", function () {
  let tableBody = document.querySelector('.ic-table__body');
  let mainWrapper = document.getElementById('ic-main__wrapper');
  const searchField = document.getElementById("ic-search-field");

  let finalData;

  // Function to fetch and process XLSX file
  async function fetchXlsxFile(filePath) {
    try {
      const response = await fetch(filePath);
      if (!response.ok) throw new Error("Failed to fetch the file");

      const data = await response.arrayBuffer(); // Read file as binary
      const workbook = XLSX.read(data, { type: "array" }); // Parse workbook

      // Assuming the first sheet is the target
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // Convert sheet to JSON (raw array of arrays format)
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // Array of arrays

      // Remove empty rows and columns
      const filteredData = jsonData
        .filter(row => row.some(cell => cell !== undefined && cell !== null && cell !== "")) // Remove empty rows
        .map(row => row.filter(cell => cell !== undefined && cell !== null && cell !== "")); // Remove empty columns

      // Convert filtered data to key-value JSON using the first row as headers
      const headers = filteredData[0]; // First row as headers
      const rows = filteredData.slice(1); // Remaining rows

      const result = rows.map(row => {
        return headers.reduce((obj, header, index) => {
          obj[header] = row[index] || ""; // Map headers to values, use empty string if value is missing
          return obj;
        }, {});
      });

      return result;
    } catch (error) {
      console.error("Error:", error);
      return null; 
    }
  }

  let showProjects = (data,dataLength)=> {
    let noResultWrapper = document.querySelector('.ic-no-result');
    if(noResultWrapper) {
      noResultWrapper.remove();
    }

    let existProjectRows = document.querySelectorAll('.ic-table__row');
    if(existProjectRows){
      existProjectRows.forEach((item)=> {
        item.remove();
      })
    }

    for(let i=0; i<dataLength; i++ ) {
      let projectRows = `<div class="ic-table__row">
                            <div class="ic-table__row-col col-sr-no"><span class="ic-table__row-col-data ">${data[i].sr_no}</span></div>
                            <div class="ic-table__row-col col-file-name"><span class="ic-table__row-col-data">${data[i].file_number}</span></div>
                            <div class="ic-table__row-col col-student-name"><span class="ic-table__row-col-data">${data[i].name_of_project_coordinator}</span></div>
                            <div class="ic-table__row-col col-final-title"><span class="ic-table__row-col-data">${data[i].final_title_of_proposal}</span></div>
                            <div class="ic-table__row-col col-pmfs"><span class="ic-table__row-col-data">${data[i].pfms_linked_account}</span></div>
                            <div class="ic-table__row-col col-debit-date"><span class="ic-table__row-col-data">${data[i].debit_date}</span></div>
                        </div>`

      tableBody.insertAdjacentHTML("beforeend", projectRows);
    }
  }

  let projectDropdown = document.getElementById("ic-project-dropdown");
  let dataFilePath = "./longitudinal.xlsx";

  fetchXlsxFile(dataFilePath).then(data => {
    finalData = data; // Assign the data after fetching
    showProjects(finalData,finalData.length)
  });
  
  projectDropdown.addEventListener("change", async () => {
    console.log(projectDropdown.value);
    if (projectDropdown.value === "vision-viksit-bharat") {
      dataFilePath = "./vision-viksit-bharat.xlsx";
    } else {
      dataFilePath = "./longitudinal.xlsx";
    }

    finalData = await fetchXlsxFile(dataFilePath);
    let existProjectRows = document.querySelectorAll('.ic-table__row');
    if(existProjectRows){
      existProjectRows.forEach((item)=> {
        item.remove();
      });
    }
    showProjects(finalData,finalData.length);
    if(searchField.value.length>0) {
      performSearch(searchField.value);

    }
  });

  const debounce = (func, delay) => {
    let timeout;
    return (...args) => {
      clearTimeout(timeout);
      timeout = setTimeout(() => func(...args), delay);
    };
  };

  // Function to perform search
  const performSearch = (query) => {
    const lowerCaseQuery = query.toLowerCase();
    const filteredResults = finalData.filter(item => 
      item.file_number.toLowerCase().includes(lowerCaseQuery)
    );

    tableBody.innerHTML = "";

    if (filteredResults.length > 0) {
      showProjects(filteredResults,filteredResults.length)
    } else {
    let noResultWrapper = document.querySelector('.ic-no-result');
      if(!noResultWrapper) {
        noResult();
      }
    }
  };

  // Attach debounced event listener to search field
  searchField.addEventListener(
    "input", 
    debounce((e) => {
      const query = e.target.value.trim();
      if (query) {
        performSearch(query);
      } else {
        showProjects(finalData,finalData.length);
      }
    }, 500) // Adjust debounce delay (500ms)
  );

  let noResult = ()=> {
    let noResultContent =`<div class="ic-no-result">
      <div class="ic-no-result__img">
        <img src="./no-result.png" alt="No Result"> 
      </div>
      <h2>No Result</h2>
      <div class="mwb-course-listing__no-result-desc">We couldn't find waht you searched for, Try searching again.</div>
  </div>`
  mainWrapper.classList.add('no-result--active');
  mainWrapper.insertAdjacentHTML('beforeend',noResultContent);
  }
});
