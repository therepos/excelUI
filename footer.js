const currentYear = new Date().getFullYear();
const org = process.env.ORG_NAME;

export default {
  style: 'dark',
  links: [],
  copyright: `
    <div class="footer-row">
      <div class="footer-left">
        <a href="https://creativecommons.org/licenses/by/4.0/" target="_blank" style="color: #ebedf0;">CC BY 4.0</a> © ${currentYear} ${org}.<br/>
        Made with Docusaurus.
      </div>
      <div class="footer-icons">
        <a href="https://github.com/therepos?tab=repositories" class="icon icon-github" target="_blank" aria-label="GitHub"></a>
        <a href="https://hub.docker.com" class="icon icon-docker" target="_blank" aria-label="Docker"></a>
        <a href="https://therepos.github.io/proxmox/" class="icon icon-proxmox" target="_blank" aria-label="Proxmox"></a>
        <a href="https://therepos.github.io/msexcel/" class="icon icon-excel" target="_blank" aria-label="Excel"></a>
        <a href="https://therepos.github.io/docusaurus/" class="icon icon-docusaurus" target="_blank" aria-label="Docusaurus"></a>            
      </div>
    </div>
  `,
};
