import React from 'react';
import Layout from '@theme-original/Layout';
import Head from '@docusaurus/Head';

export default function LayoutWrapper(props) {
  return (
    <>
      <Head>
        <link
          rel="stylesheet"
          href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined"
        />
      </Head>
      <Layout {...props} />
    </>
  );
}
