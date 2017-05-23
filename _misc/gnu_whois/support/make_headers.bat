perl -w make_as_del.pl < as_del_list > as_del.h

perl -w make_as32_del.pl < as32_del_list > as32_del.h

perl -w make_ip_del.pl < ip_del_list > ip_del.h

perl -w make_ip6_del.pl < ip6_del_list > ip6_del.h

perl -w make_tld_serv.pl < tld_serv_list > tld_serv.h
