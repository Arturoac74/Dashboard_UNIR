const colors = { 
    "Data & AI": "#38bdf8", 
    "Software Dev": "#a78bfa", 
    "Infra & Ops": "#fbbf24", 
    "Remoto": "#38bdf8", 
    "Hibrido": "#a78bfa", 
    "Presencial": "#fbbf24", 
    "IA": "#34d399" 
};

const mapColors = ['#1e293b', '#3b82f6', '#93c5fd'];
const tooltip = d3.select("#d3-tooltip");

const mapDataRAW = [
    { Pais: 'United States', Salario_Promedio_USD: 132995 }, { Pais: 'Singapore', Salario_Promedio_USD: 116955 },
    { Pais: 'Australia', Salario_Promedio_USD: 110550 }, { Pais: 'Canada', Salario_Promedio_USD: 108671 },
    { Pais: 'United Kingdom', Salario_Promedio_USD: 102826 }, { Pais: 'Netherlands', Salario_Promedio_USD: 102379 },
    { Pais: 'Germany', Salario_Promedio_USD: 100008 }, { Pais: 'Japan', Salario_Promedio_USD: 98508 },
    { Pais: 'France', Salario_Promedio_USD: 95170 }, { Pais: 'United Arab Emirates', Salario_Promedio_USD: 93508 },
    { Pais: 'Brazil', Salario_Promedio_USD: 53705 }, { Pais: 'India', Salario_Promedio_USD: 43106 }
];

const skillsCorrelationData = [
    { skill: "Python", demand: 9500, salary: 118000, category: "Data & AI" }, { skill: "PyTorch", demand: 2100, salary: 145000, category: "Data & AI" },
    { skill: "SQL", demand: 11200, salary: 95000, category: "Data & AI" }, { skill: "TypeScript", demand: 8400, salary: 112000, category: "Software Dev" },
    { skill: "JavaScript", demand: 12000, salary: 92000, category: "Software Dev" }, { skill: "Rust", demand: 1800, salary: 138000, category: "Software Dev" },
    { skill: "Go", demand: 3200, salary: 128000, category: "Software Dev" }, { skill: "Kubernetes", demand: 5400, salary: 132000, category: "Infra & Ops" },
    { skill: "Docker", demand: 8900, salary: 108000, category: "Infra & Ops" }, { skill: "Terraform", demand: 4100, salary: 125000, category: "Infra & Ops" },
    { skill: "HTML/CSS", demand: 9800, salary: 78000, category: "Software Dev" }
];

const seniorityData = [
    { level: "Junior", mediana_usd: 63542 }, { level: "Mid / Semi-Senior", mediana_usd: 79970 },
    { level: "Senior", mediana_usd: 107619 }, { level: "Lead / C-Suite", mediana_usd: 146926 }
];

const aiAdoptionData = [
    { nivel: "Nunca / Esporádico (0-1)", nivel_movil: "0-1: Esporádico", porcentaje: 7.1 }, 
    { nivel: "Ocasional (2)", nivel_movil: "2: Ocasional", porcentaje: 8.2 },
    { nivel: "Diario / Moderado (3)", nivel_movil: "3: Moderado", porcentaje: 17.8 }, 
    { nivel: "Constante (4)", nivel_movil: "4: Constante", porcentaje: 23.5 },
    { nivel: "Indispensable / Total (5)", nivel_movil: "5: Indispensable", porcentaje: 43.3 }
];

const rolesDemandData = [
    {"role": "Full-Stack Dev", "count": 12351, category: "Software Dev"}, {"role": "Back-End Dev", "count": 6453, category: "Software Dev"}, {"role": "Software Architect", "count": 2684, category: "Software Dev"}, 
    {"role": "Front-End Dev", "count": 1974, category: "Software Dev"}, {"role": "Desktop/Enterprise Dev", "count": 1919, category: "Software Dev"}, {"role": "Mobile Dev", "count": 1391, category: "Software Dev"},
    {"role": "Embedded Systems", "count": 1274, category: "Infra & Ops"}, {"role": "Engineering Manager", "count": 1068, category: "Software Dev"}, {"role": "DevOps Engineer", "count": 1053, category: "Infra & Ops"}, 
    {"role": "Data Engineer", "count": 770, category: "Data & AI"}, {"role": "AI / ML Engineer", "count": 677, category: "Data & AI"}
];

const countryRolesData = [
    {"country": "USA", "Data & AI": 170102, "Software Dev": 188450, "Infra & Ops": 246893}, {"country": "UK", "Data & AI": 143641, "Software Dev": 112122, "Infra & Ops": 112377},
    {"country": "Canadá", "Data & AI": 113483, "Software Dev": 102925, "Infra & Ops": 138892}, {"country": "Alemania", "Data & AI": 94432, "Software Dev": 87014, "Infra & Ops": 155029},
    {"country": "Francia", "Data & AI": 79854, "Software Dev": 76026, "Infra & Ops": 107816}, {"country": "India", "Data & AI": 23240, "Software Dev": 50323, "Infra & Ops": 16762}
];

const salaryData = [
    {"year": 2020, "Data & AI": 95892, "Software Dev": 99904, "Infra & Ops": 89000}, {"year": 2021, "Data & AI": 95147, "Software Dev": 101214, "Infra & Ops": 91500}, 
    {"year": 2022, "Data & AI": 95235, "Software Dev": 99248, "Infra & Ops": 92000}, {"year": 2023, "Data & AI": 94849, "Software Dev": 101259, "Infra & Ops": 94000}, 
    {"year": 2024, "Data & AI": 94935, "Software Dev": 100408, "Infra & Ops": 95000}, {"year": 2025, "Data & AI": 95069, "Software Dev": 101168, "Infra & Ops": 96500}, 
    {"year": 2026, "Data & AI": 95352, "Software Dev": 100387, "Infra & Ops": 97800}
];

const workModeData = [
    {"period": "2020-22", "Remoto": 12931, "Hibrido": 12837, "Presencial": 12845}, {"period": "2023-24", "Remoto": 8450, "Hibrido": 8449, "Presencial": 8685}, 
    {"period": "2025-26", "Remoto": 8624, "Hibrido": 8476, "Presencial": 8703}
];

function getFilterValue(id) {
    const el = document.getElementById(id);
    return el ? el.value : "ALL";
}

function exportAllDataToExcel() {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(salaryData), "Salarios_Historicos");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(skillsCorrelationData), "Skills_Correlacion");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(seniorityData), "Seniority_Vs_Salarios");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(aiAdoptionData), "Adopcion_IA");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rolesDemandData), "Roles_Demanda");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(countryRolesData), "Salarios_Rol_y_Pais");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(mapDataRAW), "Mapa_Geografico");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(workModeData), "Modalidad_Trabajo");
    XLSX.writeFile(wb, "Dashboard_TFE_Mercado_Tech_2026.xlsx");
}

google.charts.load('current', { 'packages': ['geochart'] });
google.charts.setOnLoadCallback(drawRegionsMap);
function drawRegionsMap() {
    var data = new google.visualization.DataTable();
    data.addColumn('string', 'País');
    data.addColumn('number', 'Salario Promedio (USD)');
    mapDataRAW.forEach(row => data.addRow([row.Pais, row.Salario_Promedio_USD]));
    var options = { backgroundColor: 'transparent', datalessRegionColor: '#0f172a', defaultColor: '#334155', colorAxis: { colors: mapColors }, legend: { textStyle: { color: '#94a3b8' } } };
    var chart = new google.visualization.GeoChart(document.getElementById('chart-map'));
    chart.draw(data, options);
    window.addEventListener('resize', () => chart.draw(data, options));
}

function onFilterChange() {
    d3.selectAll(".d3-chart svg").remove();
    renderAllD3();
}

function drawSkillsCorrelationChart() {
    const container = document.getElementById('chart-skills-correlation');
    if (!container) return;
    d3.select(container).selectAll("svg").remove();

    const selectedCategory = getFilterValue('categoryFilter');
    const filteredData = selectedCategory === "ALL" 
        ? skillsCorrelationData 
        : skillsCorrelationData.filter(d => d.category === selectedCategory);

    const isMobile = window.innerWidth <= 600;
    const margin = {top: 20, right: isMobile ? 15 : 30, bottom: isMobile ? 60 : 50, left: isMobile ? 45 : 60};
    const width = container.clientWidth - margin.left - margin.right;
    const height = container.clientHeight - margin.top - margin.bottom;
    
    const svg = d3.select(container).append("svg").attr("width", width + margin.left + margin.right).attr("height", height + margin.top + margin.bottom).append("g").attr("transform", `translate(${margin.left},${margin.top})`);
    const x = d3.scaleLinear().domain([0, 13000]).range([0, width]);
    const y = d3.scaleLinear().domain([70000, 150000]).range([height, 0]);

    svg.append("g").attr("class", "axis").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x).ticks(isMobile ? 4 : 8).tickFormat(d => d/1000 + "k"));
    svg.append("g").attr("class", "axis").call(d3.axisLeft(y).ticks(isMobile ? 5 : 6).tickFormat(d => "$" + d/1000 + "k"));
    svg.append("g").attr("class", "grid").call(d3.axisLeft(y).tickSize(-width).tickFormat(""));
    svg.append("g").attr("class", "grid").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x).tickSize(-height).tickFormat(""));

    svg.append("text").attr("text-anchor", "end").attr("x", width).attr("y", height + 35).attr("fill", "#94a3b8").style("font-size", isMobile ? "10px" : "12px").text("Volumen de Ofertas");
    
    svg.append("text")
        .attr("text-anchor", "middle")
        .attr("transform", "rotate(-90)")
        .attr("y", isMobile ? -32 : -45)
        .attr("x", -(height / 2))
        .attr("fill", "#94a3b8")
        .style("font-size", isMobile ? "10px" : "12px")
        .text("Salario Estimado (USD)");

    svg.append("line").attr("class", "quadrant-bg").attr("x1", x(6500)).attr("y1", 0).attr("x2", x(6500)).attr("y2", height);
    svg.append("line").attr("class", "quadrant-bg").attr("x1", 0).attr("y1", y(110000)).attr("x2", width).attr("y2", y(110000));
    
    if (!isMobile) {
        svg.append("text").attr("class", "quadrant-label").attr("x", x(12500)).attr("y", y(145000)).attr("text-anchor", "end").text("Core Cotizado");
        svg.append("text").attr("class", "quadrant-label").attr("x", x(500)).attr("y", y(145000)).attr("text-anchor", "start").text("Nichos Premium");
    }

    const node = svg.append("g").selectAll("g").data(filteredData).join("g").attr("transform", d => `translate(${x(d.demand)},${y(d.salary)})`);

    node.append("circle").attr("r", 0).attr("fill", d => colors[d.category]).attr("opacity", 0.7).attr("stroke", "#0f172a").attr("stroke-width", 2)
        .on("mouseover", function(e, d) { if(!isMobile) { d3.select(this).attr("opacity", 1).attr("r", 15); tooltip.style("display", "block").style("opacity", 1).html(`<strong>${d.skill}</strong>Demanda: ${d.demand.toLocaleString()} ofertas<br>Salario Base: $${d.salary.toLocaleString()} USD<br><em>Clasificación: ${d.category}</em>`); }})
        .on("mousemove", e => { if(!isMobile) tooltip.style("left", (e.pageX + 15) + "px").style("top", (e.pageY - 28) + "px"); })
        .on("mouseleave", function() { if(!isMobile) { d3.select(this).attr("opacity", 0.7).attr("r", isMobile ? 8 : 12); tooltip.style("opacity", 0).style("display", "none"); }})
        .transition().duration(1000).delay((d, i) => i * 100).attr("r", isMobile ? 8 : 12);

    node.append("text").attr("dy", isMobile ? -12 : -16).attr("text-anchor", "middle").style("font-size", isMobile ? "9px" : "11px").style("fill", "#e2e8f0").style("font-weight", "500").text(d => d.skill)
        .style("opacity", 0).transition().duration(1000).delay((d, i) => (i * 100) + 500).style("opacity", 1);
}

function drawSeniorityChart() {
    const container = document.getElementById('chart-seniority');
    if (!container) return;
    d3.select(container).selectAll("svg").remove();

    const isMobile = window.innerWidth <= 600;
    const margin = {top: 20, right: 20, bottom: 40, left: isMobile ? 45 : 55};
    const width = container.clientWidth - margin.left - margin.right;
    const height = container.clientHeight - margin.top - margin.bottom;
    const svg = d3.select(container).append("svg").attr("width", width + margin.left + margin.right).attr("height", height + margin.top + margin.bottom).append("g").attr("transform", `translate(${margin.left},${margin.top})`);
    
    const x = d3.scaleBand().domain(seniorityData.map(d => d.level)).range([0, width]).padding(0.3);
    const y = d3.scaleLinear().domain([0, 160000]).range([height, 0]);

    svg.append("g").attr("class", "axis").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x)).selectAll("text").style("font-size", isMobile ? "9px" : "11px");
    svg.append("g").attr("class", "axis").call(d3.axisLeft(y).ticks(5).tickFormat(d => "$" + d/1000 + "k"));
    svg.append("g").attr("class", "grid").call(d3.axisLeft(y).tickSize(-width).tickFormat(""));

    svg.selectAll("rect").data(seniorityData).join("rect")
        .attr("x", d => x(d.level)).attr("width", x.bandwidth())
        .attr("fill", "#a78bfa").attr("rx", 4)
        .attr("y", height).attr("height", 0)
        .on("mouseover", function(e, d) { if(!isMobile) { d3.select(this).attr("fill", "#c4b5fd"); tooltip.style("display", "block").style("opacity", 1).html(`<strong>${d.level}</strong>Mediana Salarial: $${d.mediana_usd.toLocaleString()} USD`); }})
        .on("mousemove", e => { if(!isMobile) tooltip.style("left", (e.pageX + 15) + "px").style("top", (e.pageY - 28) + "px"); })
        .on("mouseleave", function() { if(!isMobile) { d3.select(this).attr("fill", "#a78bfa"); tooltip.style("opacity", 0).style("display", "none"); }})
        .transition().duration(1000).delay((d, i) => i * 150)
        .attr("y", d => y(d.mediana_usd)).attr("height", d => height - y(d.mediana_usd));
}

function drawAIAdoptionChart() {
    const container = document.getElementById('chart-ai-adoption');
    if (!container) return;
    d3.select(container).selectAll("svg").remove();

    const isMobile = window.innerWidth <= 600;
    const margin = {top: 20, right: 15, bottom: 40, left: isMobile ? 100 : 135};
    const width = container.clientWidth - margin.left - margin.right;
    const height = container.clientHeight - margin.top - margin.bottom;
    const svg = d3.select(container).append("svg").attr("width", width + margin.left + margin.right).attr("height", height + margin.top + margin.bottom).append("g").attr("transform", `translate(${margin.left},${margin.top})`);
    
    const x = d3.scaleLinear().domain([0, 50]).range([0, width]);
    const y = d3.scaleBand().domain(aiAdoptionData.map(d => isMobile ? d.nivel_movil : d.nivel)).range([height, 0]).padding(0.25);

    svg.append("g").attr("class", "axis").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x).ticks(5).tickFormat(d => d + "%"));
    svg.append("g").attr("class", "axis").call(d3.axisLeft(y)).selectAll("text").style("font-size", isMobile ? "9px" : "11px");
    svg.append("g").attr("class", "grid").call(d3.axisBottom(x).tickSize(-height).tickFormat(""));

    svg.selectAll("rect").data(aiAdoptionData).join("rect")
        .attr("x", 0).attr("y", d => y(isMobile ? d.nivel_movil : d.nivel)).attr("height", y.bandwidth())
        .attr("fill", "#34d399").attr("rx", 3).attr("width", 0)
        .on("mouseover", function(e, d) { if(!isMobile) { d3.select(this).attr("fill", "#6ee7b7"); tooltip.style("display", "block").style("opacity", 1).html(`<strong>${d.nivel}</strong>Participación: ${d.porcentaje}% de desarrolladores`); }})
        .on("mousemove", e => { if(!isMobile) tooltip.style("left", (e.pageX + 15) + "px").style("top", (e.pageY - 28) + "px"); })
        .on("mouseleave", function() { if(!isMobile) { d3.select(this).attr("fill", "#34d399"); tooltip.style("opacity", 0).style("display", "none"); }})
        .transition().duration(1000).delay((d, i) => i * 100)
        .attr("width", d => x(d.porcentaje));
}

function drawRolesDemandChart() {
    const container = document.getElementById('chart-roles-demand');
    if (!container) return;
    d3.select(container).selectAll("svg").remove();

    const selectedCategory = getFilterValue('categoryFilter');
    const filteredRoles = selectedCategory === "ALL" 
        ? rolesDemandData 
        : rolesDemandData.filter(d => d.category === selectedCategory);

    const isMobile = window.innerWidth <= 600;
    const margin = {top: 10, right: 30, bottom: 20, left: isMobile ? 110 : 160};
    const width = container.clientWidth - margin.left - margin.right;
    const height = container.clientHeight - margin.top - margin.bottom;
    const svg = d3.select(container).append("svg").attr("width", width + margin.left + margin.right).attr("height", height + margin.top + margin.bottom).append("g").attr("transform", `translate(${margin.left},${margin.top})`);
    
    filteredRoles.sort((a, b) => a.count - b.count); 
    const x = d3.scaleLinear().domain([0, d3.max(rolesDemandData, d => d.count) + 1500]).range([0, width]);
    const y = d3.scaleBand().domain(filteredRoles.map(d => d.role)).range([height, 0]).padding(0.25);
    
    svg.append("g").attr("class", "axis").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x).ticks(isMobile ? 3 : 8));
    svg.append("g").attr("class", "axis").call(d3.axisLeft(y)).selectAll("text").style("font-size", isMobile ? "9px" : "11px");
    
    svg.selectAll("rect").data(filteredRoles).join("rect").attr("x", 0).attr("y", d => y(d.role)).attr("height", y.bandwidth()).attr("fill", "#34d399").attr("rx", 3).attr("width", 0)
        .on("mouseover", function(e, d) { if(!isMobile) { d3.select(this).attr("fill", "#6ee7b7"); tooltip.style("display", "block").style("opacity", 1).html(`<strong>${d.role}</strong>Participación: ${d.count.toLocaleString()} de vacantes`); }})
        .on("mousemove", e => { if(!isMobile) tooltip.style("left", (e.pageX + 15) + "px").style("top", (e.pageY - 28) + "px"); })
        .on("mouseleave", function() { if(!isMobile) { d3.select(this).attr("fill", "#34d399"); tooltip.style("opacity", 0).style("display", "none"); }})
        .transition().duration(1000).delay((d, i) => i * 80).attr("width", d => x(d.count));
}

function drawCountryRolesChart() {
    const container = document.getElementById('chart-country-roles');
    if (!container) return;
    d3.select(container).selectAll("svg").remove();

    const selectedCategory = getFilterValue('categoryFilter');
    const subgroups = selectedCategory === "ALL" ? ["Data & AI", "Software Dev", "Infra & Ops"] : [selectedCategory];

    const isMobile = window.innerWidth <= 600;
    const margin = {top: isMobile ? 40 : 20, right: isMobile ? 10 : 120, bottom: 40, left: isMobile ? 40 : 60};
    const width = container.clientWidth - margin.left - margin.right;
    const height = container.clientHeight - margin.top - margin.bottom;
    const svg = d3.select(container).append("svg").attr("width", width + margin.left + margin.right).attr("height", height + margin.top + margin.bottom).append("g").attr("transform", `translate(${margin.left},${margin.top})`);
    
    const groups = countryRolesData.map(d => d.country);
    const x = d3.scaleBand().domain(groups).range([0, width]).padding([0.15]);
    const xSubgroup = d3.scaleBand().domain(subgroups).range([0, x.bandwidth()]).padding([0.05]);
    const y = d3.scaleLinear().domain([0, 260000]).range([height, 0]);
    
    svg.append("g").attr("class", "axis").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x)).selectAll("text").style("font-size", isMobile ? "9px" : "11px");
    svg.append("g").attr("class", "axis").call(d3.axisLeft(y).ticks(isMobile ? 4 : 6).tickFormat(d => "$" + d / 1000 + "k"));
    svg.append("g").attr("class", "grid").call(d3.axisLeft(y).tickSize(-width).tickFormat(""));
    
    svg.append("g").selectAll("g").data(countryRolesData).join("g").attr("transform", d => `translate(${x(d.country)}, 0)`)
        .selectAll("rect").data(d => subgroups.map(key => ({country: d.country, key: key, value: d[key]})))
        .join("rect").attr("x", d => xSubgroup(d.key)).attr("fill", d => colors[d.key]).attr("rx", 2).attr("width", xSubgroup.bandwidth())
        .attr("y", height).attr("height", 0)
        .on("mouseover", function(e, d) { if(!isMobile) { d3.select(this).style("opacity", 0.75); tooltip.style("display", "block").style("opacity", 1).html(`<b>País:</b> ${d.country}<br/><b>Categoría:</b> ${d.key}<br/><b>Salario:</b> $${d.value.toLocaleString()}`); }})
        .on("mousemove", e => { if(!isMobile) tooltip.style("left", (e.pageX + 15) + "px").style("top", (e.pageY - 28) + "px"); })
        .on("mouseleave", function() { if(!isMobile) { d3.select(this).style("opacity", 1); tooltip.style("opacity", 0).style("display", "none"); }})
        .transition().duration(1000).delay((d, i) => i * 150).attr("y", d => y(d.value)).attr("height", d => height - y(d.value));
    
    const legend = svg.append("g").attr("transform", isMobile ? `translate(0, -30)` : `translate(${width + 10}, 0)`);
    subgroups.forEach((key, i) => {
        if (isMobile) {
            legend.append("rect").attr("x", i * (width/subgroups.length)).attr("y", 0).attr("width", 10).attr("height", 10).attr("fill", colors[key]).attr("rx", 2);
            legend.append("text").attr("x", i * (width/subgroups.length) + 15).attr("y", 9).style("fill", "#cbd5e1").style("font-size", "9px").text(key.replace(' & ', '&'));
        } else {
            legend.append("rect").attr("y", i * 20).attr("width", 12).attr("height", 12).attr("fill", colors[key]).attr("rx", 2);
            legend.append("text").attr("x", 20).attr("y", i * 20 + 10).style("fill", "#cbd5e1").style("font-size", "12px").text(key);
        }
    });
}

function drawSalaryChart() {
    const container = document.getElementById('chart-salary');
    if (!container) return;
    d3.select(container).selectAll("svg").remove();

    const selectedCategory = getFilterValue('categoryFilter');
    const categoriesToDraw = selectedCategory === "ALL" ? ["Data & AI", "Software Dev", "Infra & Ops"] : [selectedCategory];

    const isMobile = window.innerWidth <= 600;
    const margin = {top: isMobile ? 35 : 20, right: isMobile ? 15 : 100, bottom: 30, left: isMobile ? 45 : 50};
    const width = container.clientWidth - margin.left - margin.right;
    const height = container.clientHeight - margin.top - margin.bottom;

    const svg = d3.select(container)
        .append("svg")
        .attr("width", width + margin.left + margin.right)
        .attr("height", height + margin.top + margin.bottom)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleLinear().domain(d3.extent(salaryData, d => d.year)).range([0, width]);
    const y = d3.scaleLinear().domain([85000, 105000]).range([height, 0]);
    
    svg.append("g").attr("class", "axis").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x).tickFormat(d3.format("d")).ticks(isMobile ? 4 : 7));
    svg.append("g").attr("class", "axis").call(d3.axisLeft(y).ticks(isMobile ? 4 : 5).tickFormat(d => "$" + d/1000 + "k"));
    svg.append("g").attr("class", "grid").call(d3.axisLeft(y).tickSize(-width).tickFormat(""));

    if (isMobile) {
        const legend = svg.append("g").attr("transform", `translate(0, -25)`);
        categoriesToDraw.forEach((cat, i) => {
            legend.append("rect").attr("x", i * (width/categoriesToDraw.length)).attr("y", 0).attr("width", 10).attr("height", 10).attr("fill", colors[cat]).attr("rx", 2);
            legend.append("text").attr("x", i * (width/categoriesToDraw.length) + 14).attr("y", 9).style("fill", "#cbd5e1").style("font-size", "9px").text(cat.replace(' & ', '&'));
        });
    }

    categoriesToDraw.forEach((cat, index) => {
        const lineGenerator = d3.line()
            .curve(d3.curveMonotoneX)
            .x(d => x(d.year))
            .y(d => y(d[cat]));

        const path = svg.append("path")
            .datum(salaryData)
            .attr("fill", "none")
            .attr("stroke", colors[cat])
            .attr("stroke-width", 3)
            .attr("d", lineGenerator);

        const totalLength = path.node().getTotalLength();
        path.attr("stroke-dasharray", totalLength + " " + totalLength)
            .attr("stroke-dashoffset", totalLength)
            .transition()
            .duration(2000)
            .delay(index * 200)
            .ease(d3.easeCubicInOut)
            .attr("stroke-dashoffset", 0);

        const className = `dot-${cat.replace(/[^a-zA-Z0-9]/g, '')}`;

        svg.selectAll(`.${className}`)
            .data(salaryData)
            .enter()
            .append("circle")
            .attr("class", className)
            .attr("cx", d => x(d.year))
            .attr("cy", d => y(d[cat]))
            .attr("r", 5) 
            .attr("fill", colors[cat])
            .attr("stroke", "#ffffff") 
            .attr("stroke-width", 1.5)
            .style("opacity", 1) 
            .style("cursor", "pointer")
            .on("mouseover", function(e, d) {
                if (!isMobile) {
                    d3.select(this)
                        .transition().duration(150)
                        .attr("r", 8);

                    tooltip
                        .style("display", "block")
                        .style("opacity", 1)
                        .style("left", (e.pageX + 15) + "px")
                        .style("top", (e.pageY - 28) + "px")
                        .html(`<strong>${cat}</strong> (${d.year})<br/>Salario: $${d3.format(",")(d[cat])}`);
                }
            })
            .on("mousemove", e => {
                if (!isMobile) {
                    tooltip
                        .style("left", (e.pageX + 15) + "px")
                        .style("top", (e.pageY - 28) + "px");
                }
            })
            .on("mouseleave", function() {
                if (!isMobile) {
                    d3.select(this)
                        .transition().duration(150)
                        .attr("r", 5)
                        .attr("stroke-width", 1.5);

                    tooltip
                        .style("opacity", 0)
                        .style("display", "none");
                }
            });
    });
}

function drawWorkModeChart() {
    const container = document.getElementById('chart-workmode');
    if (!container) return;
    d3.select(container).selectAll("svg").remove();

    const selectedWorkMode = getFilterValue('workModeFilter');
    const subgroups = selectedWorkMode === "ALL" ? ["Remoto", "Hibrido", "Presencial"] : [selectedWorkMode];

    const isMobile = window.innerWidth <= 600;
    const margin = {top: isMobile ? 35 : 20, right: isMobile ? 10 : 20, bottom: 40, left: isMobile ? 35 : 50};
    const width = container.clientWidth - margin.left - margin.right;
    const height = container.clientHeight - margin.top - margin.bottom;
    const svg = d3.select(container).append("svg").attr("width", width + margin.left + margin.right).attr("height", height + margin.top + margin.bottom).append("g").attr("transform", `translate(${margin.left},${margin.top})`);
    
    const groups = workModeData.map(d => d.period);
    const x = d3.scaleBand().domain(groups).range([0, width]).padding([0.2]);
    const xSubgroup = d3.scaleBand().domain(subgroups).range([0, x.bandwidth()]).padding([0.05]);
    const y = d3.scaleLinear().domain([0, 14000]).range([height, 0]);
    
    svg.append("g").attr("class", "axis").attr("transform", `translate(0,${height})`).call(d3.axisBottom(x)).selectAll("text").style("font-size", isMobile ? "10px" : "11px");
    svg.append("g").attr("class", "axis").call(d3.axisLeft(y).ticks(4).tickFormat(d => d/1000 + "k"));
    
    svg.append("g").selectAll("g").data(workModeData).join("g").attr("transform", d => `translate(${x(d.period)}, 0)`).selectAll("rect").data(d => subgroups.map(key => ({period: d.period, key: key, value: d[key]})))
        .join("rect").attr("x", d => xSubgroup(d.key)).attr("fill", d => colors[d.key]).attr("rx", 2).attr("width", xSubgroup.bandwidth()).attr("y", height).attr("height", 0)
        .on("mouseover", function(e, d) { if(!isMobile) { d3.select(this).style("opacity", 0.75); tooltip.style("display", "block").style("opacity", 1).html(`<b>Modalidad:</b> ${d.key}<br/><b>Período:</b> ${d.period}<br/><b>Salario:</b> $${d.value.toLocaleString()}`); }})
        .on("mousemove", e => { if(!isMobile) tooltip.style("left", (e.pageX + 15) + "px").style("top", (e.pageY - 28) + "px"); })
        .on("mouseleave", function() { if(!isMobile) { d3.select(this).style("opacity", 1); tooltip.style("opacity", 0).style("display", "none"); }})
        .transition().duration(1000).delay((d, i) => i * 150).attr("y", d => y(d.value)).attr("height", d => height - y(d.value));
        
    const legend = svg.append("g").attr("transform", isMobile ? `translate(0, -25)` : `translate(${width - 80}, -10)`);
    subgroups.forEach((key, i) => {
        if(isMobile) {
            legend.append("rect").attr("x", i * (width/subgroups.length)).attr("y", 0).attr("width", 10).attr("height", 10).attr("fill", colors[key]).attr("rx", 2);
            legend.append("text").attr("x", i * (width/subgroups.length) + 14).attr("y", 9).style("fill", "#cbd5e1").style("font-size", "9px").text(key);
        } else {
            legend.append("rect").attr("y", i * 15).attr("width", 10).attr("height", 10).attr("fill", colors[key]).attr("rx", 2);
            legend.append("text").attr("x", 15).attr("y", i * 15 + 9).style("fill", "#cbd5e1").style("font-size", "11px").text(key);
        }
    });
}

function renderAllD3() { 
    drawSkillsCorrelationChart(); 
    drawSeniorityChart(); 
    drawAIAdoptionChart(); 
    drawRolesDemandChart(); 
    drawCountryRolesChart(); 
    drawSalaryChart(); 
    drawWorkModeChart(); 
}

renderAllD3();
window.addEventListener("resize", () => { d3.selectAll(".d3-chart svg").remove(); renderAllD3(); });

function switchTab(tabId, element) {
    
    document.querySelectorAll('.nav-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));

   
    element.classList.add('active');
    document.getElementById('tab-' + tabId).classList.add('active');

    
    setTimeout(() => {
        window.dispatchEvent(new Event('resize'));
    }, 50);
}