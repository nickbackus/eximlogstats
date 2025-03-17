#!/bin/bash
/bin/cp /var/log/exim_mainlog [fullpath]/eximlog
/bin/chown [user] [fullpath]/eximlog
/bin/chgrp [user] [fullpath]/eximlog