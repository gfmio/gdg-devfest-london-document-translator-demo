import React, { Fragment, FC } from "react";
import AppBar from "@material-ui/core/AppBar";
import Toolbar from "@material-ui/core/Toolbar";
import Typography from "@material-ui/core/Typography";
import Head from "next/head";

const Layout: FC = ({ children }) => (
  <Fragment>
    <Head>
      <title>Google Cloud Translation Document Translator</title>
      <link rel="icon" href="/favicon.ico" />
      <link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Roboto:300,400,500,700&display=swap" />
    </Head>
    <AppBar position="static">
      <Toolbar>
        <Typography variant="h6" component="h1">
          Google Cloud Translation Document Translator
        </Typography>
      </Toolbar>
    </AppBar>
    <div>{children}</div>
    <style jsx global>
      {`
        body {
          margin: 0;
        }
      `}
    </style>
  </Fragment>
);

export default Layout;
