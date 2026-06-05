
// 文档交互功能
document.addEventListener('DOMContentLoaded', function() {
    // 返回顶部按钮
    const backToTop = document.createElement('button');
    backToTop.textContent = '↑';
    backToTop.className = 'back-to-top';
    backToTop.style.cssText = `
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: #4f46e5;
        color: white;
        border: none;
        border-radius: 50%;
        width: 50px;
        height: 50px;
        font-size: 20px;
        cursor: pointer;
        display: none;
        z-index: 1000;
    `;
    document.body.appendChild(backToTop);
    
    // 监听滚动显示/隐藏返回顶部按钮
    window.addEventListener('scroll', function() {
        if (window.pageYOffset > 300) {
            backToTop.style.display = 'block';
        } else {
            backToTop.style.display = 'none';
        }
    });
    
    backToTop.addEventListener('click', function() {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });
    
    // 代码块复制按钮
    document.querySelectorAll('pre').forEach(function(pre) {
        const copyButton = document.createElement('button');
        copyButton.textContent = '复制';
        copyButton.className = 'copy-button';
        copyButton.style.cssText = `
            position: absolute;
            top: 10px;
            right: 10px;
            background: #6b7280;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 5px 10px;
            font-size: 12px;
            cursor: pointer;
        `;
        
        pre.style.position = 'relative';
        pre.appendChild(copyButton);
        
        copyButton.addEventListener('click', function() {
            const code = pre.textContent;
            navigator.clipboard.writeText(code).then(function() {
                copyButton.textContent = '已复制';
                setTimeout(function() {
                    copyButton.textContent = '复制';
                }, 2000);
            });
        });
    });
    
    // 添加导航菜单
    const headers = document.querySelectorAll('h1, h2, h3');
    const nav = document.createElement('nav');
    nav.className = 'nav-menu';
    nav.innerHTML = '<h3>快速导航</h3><ul></ul>';
    
    const navList = nav.querySelector('ul');
    
    headers.forEach(function(header, index) {
        if (header.tagName === 'H1' || header.tagName === 'H2') {
            const li = document.createElement('li');
            li.innerHTML = `<a href="#${header.id}">${header.textContent}</a>`;
            navList.appendChild(li);
        }
    });
    
    // 找到第一个H2后插入导航
    const firstH2 = document.querySelector('h2');
    if (firstH2) {
        firstH2.parentNode.insertBefore(nav, firstH2);
    }
});

// 搜索高亮功能
function highlightSearchTerm(term) {
    const content = document.querySelector('.readable-content');
    const text = content.innerHTML;
    const regex = new RegExp(`(${term})`, 'gi');
    content.innerHTML = text.replace(regex, '<mark>$1</mark>');
}
