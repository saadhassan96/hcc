/* HCC Group site behavior: theme switch, mobile menu, news and publication lists.
   Everything works without JavaScript; this only adds filtering and collapsing. */
(function () {
  var root = document.documentElement;
  root.classList.remove('no-js');
  root.classList.add('js');

  /* ---------- theme ---------- */
  var mq = window.matchMedia('(prefers-color-scheme: dark)');
  function isDark() {
    var t = root.dataset.theme;
    return t ? t === 'dark' : mq.matches;
  }
  function syncTheme() {
    var dark = isDark();
    root.classList.toggle('dark-ui', dark);
    var btn = document.querySelector('.theme-btn');
    if (btn) btn.setAttribute('aria-label', dark ? 'Use light theme' : 'Use dark theme');
  }
  if (mq.addEventListener) mq.addEventListener('change', syncTheme);

  document.addEventListener('DOMContentLoaded', function () {
    syncTheme();
    var themeBtn = document.querySelector('.theme-btn');
    themeBtn.addEventListener('click', function () {
      var next = isDark() ? 'light' : 'dark';
      root.dataset.theme = next;
      try { localStorage.setItem('hcc-theme', next); } catch (e) {}
      syncTheme();
    });

    /* ---------- mobile menu ---------- */
    var menuBtn = document.querySelector('.menu-btn');
    var nav = document.getElementById('site-nav');
    function closeMenu() { nav.classList.remove('open'); menuBtn.setAttribute('aria-expanded', 'false'); }
    menuBtn.addEventListener('click', function () {
      var open = !nav.classList.contains('open');
      nav.classList.toggle('open', open);
      menuBtn.setAttribute('aria-expanded', String(open));
    });
    nav.addEventListener('click', function (e) { if (e.target.closest('a')) closeMenu(); });
    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && nav.classList.contains('open')) { closeMenu(); menuBtn.focus(); }
    });

    /* ---------- current section in nav ---------- */
    var links = Array.prototype.slice.call(nav.querySelectorAll('a[href^="#"]'));
    if ('IntersectionObserver' in window) {
      var obs = new IntersectionObserver(function (entries) {
        entries.forEach(function (en) {
          if (!en.isIntersecting) return;
          links.forEach(function (a) {
            if (a.getAttribute('href') === '#' + en.target.id) a.setAttribute('aria-current', 'true');
            else a.removeAttribute('aria-current');
          });
        });
      }, { rootMargin: '-45% 0px -50% 0px' });
      links.forEach(function (a) {
        var s = document.querySelector(a.getAttribute('href'));
        if (s && !a.classList.contains('join-link')) obs.observe(s);
      });
      var top = document.getElementById('top');
      if (top) obs.observe(top); // no nav item for the hero, so this clears the highlight
    }

    /* ---------- news: show all / fewer ---------- */
    document.querySelectorAll('[data-toggle-list]').forEach(function (btn) {
      var list = document.getElementById(btn.dataset.toggleList);
      btn.setAttribute('aria-controls', list.id);
      btn.addEventListener('click', function () {
        var open = !list.classList.contains('expanded');
        list.classList.toggle('expanded', open);
        btn.setAttribute('aria-expanded', String(open));
        btn.textContent = open ? btn.dataset.less : btn.dataset.more;
        if (!open) list.scrollIntoView({ block: 'nearest' });
      });
    });

    /* ---------- publications: filters and collapse ---------- */
    var list = document.getElementById('pub-list');
    if (!list) return;
    var pubs = Array.prototype.slice.call(list.querySelectorAll('.pub'));
    var years = Array.prototype.slice.call(list.querySelectorAll('.years > li, #pub-list > li'));
    var leadSel = document.getElementById('lead-filter');
    var projectSel = document.getElementById('project-filter');
    var status = document.getElementById('filter-status');
    var toggle = document.getElementById('pubs-toggle');
    var LIMIT = 12;
    var state = { lead: '', project: '', expanded: false };

    function apply(announce) {
      var filtered = state.lead || state.project;
      var shown = 0, matches = 0;
      pubs.forEach(function (p) {
        var ok = (!state.lead || (' ' + p.dataset.lead + ' ').indexOf(' ' + state.lead + ' ') > -1) &&
                 (!state.project || p.dataset.project === state.project);
        if (ok) matches++;
        var visible = ok && (filtered || state.expanded || matches <= LIMIT);
        p.hidden = !visible;
        if (visible) shown++;
      });
      years.forEach(function (y) { y.hidden = !y.querySelector('.pub:not([hidden])'); });
      toggle.hidden = !!filtered || pubs.length <= LIMIT;
      toggle.setAttribute('aria-expanded', String(state.expanded));
      toggle.textContent = state.expanded ? 'Show fewer publications' : 'Show all ' + pubs.length + ' publications';
      var msg;
      if (filtered) msg = matches === 1 ? '1 publication matches.' : matches + ' publications match.';
      else msg = state.expanded ? 'Showing all ' + pubs.length + ' publications.' : 'Showing the ' + shown + ' most recent of ' + pubs.length + ' publications.';
      if (filtered && !matches) msg = 'No publications match. Choose another co-director or area.';
      status.textContent = announce || filtered ? msg : '';
      if (!filtered && !announce) status.textContent = '';
    }
    leadSel.addEventListener('change', function () { state.lead = leadSel.value; apply(true); });
    projectSel.addEventListener('change', function () { state.project = projectSel.value; apply(true); });
    toggle.addEventListener('click', function () {
      state.expanded = !state.expanded;
      apply(true);
      if (!state.expanded) document.getElementById('all-pubs').scrollIntoView({ block: 'start' });
    });
    document.querySelectorAll('[data-filter-project]').forEach(function (a) {
      a.addEventListener('click', function () {
        state.project = projectSel.value = a.dataset.filterProject;
        state.lead = leadSel.value = '';
        apply(true);
      });
    });
    apply(false);
  });
})();
