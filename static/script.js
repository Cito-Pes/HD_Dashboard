// 그래프 페이지 순환 로직
const pages = ['month_chart', 'week_chart', 'day_chart', 'pie_chart'];
let currentPage = 0;

function switchPage() {
    pages.forEach((id, index) => {
        document.getElementById(id).style.display = (index === currentPage) ? 'block' : 'none';
    });
    currentPage = (currentPage + 1) % pages.length;
}

setInterval(switchPage, 20000); // 20초 간격
switchPage();