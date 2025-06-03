import { themes as prismThemes } from 'prism-react-renderer';

const currentYear = new Date().getFullYear();

export default {
  title: 'Excel',
  tagline: 'Excel',
  url: 'https://therepos.github.io',
  baseUrl: '/msexcel/',
  organizationName: 'therepos',
  projectName: 'msexcel',
  deploymentBranch: 'gh-pages',
  trailingSlash: false,

  presets: [
    [
      '@docusaurus/preset-classic',
      {
        docs: {
          path: 'docs',
          routeBasePath: '/',
          sidebarPath: './sidebars.js',
          showLastUpdateTime: true,
          sidebarCollapsible: true,
          editUrl: 'https://github.com/therepos/msexcel/edit/main/',
        },
        theme: {
          customCss: './src/css/styles.css',
        },
      },
    ],
  ],

  themeConfig: {
    navbar: {
      title: 'Excel',
      items: [
        {
          type: 'search',
          position: 'right',
        },
        {
          href: 'https://github.com/therepos/msexcel',
          position: 'right',
          className: 'header-github-link',
          'aria-label': 'GitHub repository',
        },
      ],
    },
    docs: {
      sidebar: {
        hideable: true,
        autoCollapseCategories: true,
      },
    },
    prism: {
      theme: prismThemes.github,
      additionalLanguages: ['git'],
    },
    footer: {
      style: 'dark',
      links: [],
      copyright: `
        <div class="footer-row">
          <div class="footer-left">
            <a href="https://creativecommons.org/licenses/by/4.0/" target="_blank" style="color: #ebedf0;">CC BY 4.0</a> © ${currentYear} therepos.<br/>
            Made with Docusaurus.
          </div>
          <div class="footer-icons">
            <a href="https://github.com" class="icon icon-github" target="_blank" aria-label="GitHub"></a>
            <a href="https://hub.docker.com" class="icon icon-docker" target="_blank" aria-label="Docker"></a>
          </div>
        </div>
      `,
    },
  },
};
