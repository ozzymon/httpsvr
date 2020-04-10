#include <stdio.h>
#include <stdlib.h>
#include <unistd.h>
#include <string.h>
#include <sys/ioctl.h>
#include <sys/socket.h>
#include <netinet/in.h>
#include <arpa/inet.h>
#include <time.h>
#include <signal.h>

#define PORT 8080
#define BACKLOG 10
#define RNGSIZE 10
#define MCONN 2
#define FDSSIZE (MCONN * RNGSIZE + 1)
#define LEN_REQ 1024
#define LEN_RES 1536
#define LEN_RFC1123_TIME 30
#define TIMEOUT 2
#define LEN_PST 512
#define LEN_HTML 1024
#define RECORD 128

struct io {
	int fd;
	int flg_i;
	int flg_o;
	int rcvd;
	int hbdy;
	int sent;
	time_t t_res;
	char req[LEN_REQ];
	char res[LEN_RES];
	struct sockaddr_in6 addr;
};

struct fd {
	int fd;
	time_t t_acpt;
	struct sockaddr_in6 addr;
	struct io *io;
};

struct record {
	int flg;
	time_t t_rec;
	struct in6_addr addr;
	char post[LEN_PST];
};

static const char h200[] = "HTTP/1.0 200 OK";
static const char h404[] = "HTTP/1.0 404 Not Found";
static const char hsvr[] = "Server: Prototype webserver";
static const char hctplain[] = "Content-Type: text/plain; charset=utf-8";

static const char *DAY_NAMES[] = {"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"};
static const char *MONTH_NAMES[] = {"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};

enum method{HEAD, GET, POST, HCSV, GCSV, N404};

int gshutdown = 1;
char html[LEN_HTML];

int recv_chk(int sockfd);
int recv_req(int sockfd, char *buffer, int *rcvd);
int crlf2(char *buffer);
int content_length(char *buffer, int hbdy, int rcvd);
int send_res(int sockfd, int len, char *buffer);
void init_s_io(struct io *io);
void init_s_fd(struct fd *fd);
int io_http(struct io *io, struct record *rec);
void proc_post(char *dst, char *src);
int proc_csv(char *buffer, size_t n, struct record *rec);
int save_rec(struct record *rec, int next);
int Rfc1123_DateTime(char *buffer, size_t n, time_t *t);
void fatal_error(const char *message);

void handle_shutdown(int signal) {
	fprintf(stderr, "\nStopping...\n");
	gshutdown = 0;
}

int main(){
	struct sockaddr_in6 p_addr, c_addr;
	struct fd fd[FDSSIZE];
	struct io io[RNGSIZE];
	struct record rec[RECORD];
	struct timeval zero;

	char date[LEN_RFC1123_TIME];
	int p_sock, c_sock, usedfds, nextrb, nextrec, maxfd, i, j, yes = 1;
	FILE *fp;
	fd_set fds, fds_e;
	time_t t;

	size_t sin6_size = sizeof(struct sockaddr_in6);

	//get html
	fp = fopen("./index.html", "r");
	if (fp == NULL)
		fatal_error("Opening HTML File");

	i = 0;
	while (!feof(fp) && i < LEN_HTML) {
		html[i] = fgetc(fp);
		i++;
	}

	if (i == 0)
		fatal_error("Opening HTML File");

	html[i - 1] = 0x00;
	fclose(fp);

	p_sock = socket(PF_INET6, SOCK_STREAM, 0);
	if (p_sock == -1)
		fatal_error("Creating socket");

	if (setsockopt(p_sock, SOL_SOCKET, SO_REUSEADDR, &yes, sizeof(int)) == -1)
		fatal_error("Setting SO_REUSEADDR");

	if (ioctl(p_sock, FIONBIO, &yes) == -1)
		fatal_error("Setting FIONBIO");

	memset(&p_addr, 0x00, sin6_size);
	p_addr.sin6_family = AF_INET6;
	p_addr.sin6_port = htons(PORT);
	memcpy(&(p_addr.sin6_addr), &in6addr_any, sizeof(struct in6_addr));

	if (bind(p_sock, (struct sockaddr *)&p_addr, (socklen_t)sin6_size) == -1)
		fatal_error("Binding to socket");

	if (listen(p_sock, BACKLOG) == -1)
		fatal_error("Listening on socket");

	t = time(NULL);
	fprintf(stderr, "listening:%d %lu\n", p_sock, t);
	usedfds = 1;
	nextrb = 0;
	nextrec = 0;
	memset(fd, 0x00, sizeof(struct fd) * FDSSIZE);
	fd[0].fd = p_sock;
	for(i = 1; i < FDSSIZE; i++)
		fd[i].fd = -1;
	memset(io, 0x00, sizeof(struct io) * RNGSIZE);
	memset(rec, 0x00, sizeof(struct record) * RECORD);

	//sentinel
	rec[0].flg = -1;

	signal(SIGTSTP, handle_shutdown);
	signal(SIGPIPE, SIG_IGN);

	while(gshutdown){
	//init timeout fd
		t = time(NULL);
		for(i = 1; i < FDSSIZE; i++){
			if ((fd[i].t_acpt != 0) && (t - fd[i].t_acpt) >= TIMEOUT){
				fprintf(stderr, "timeout:%d %lu\n", fd[i].fd, t);
				close(fd[i].fd);
				if (fd[i].io != 0x00)
					init_s_io(fd[i].io);
				init_s_fd(fd + i);
				usedfds--;
			}
		}

	//init fd_set
		FD_ZERO(&fds);
		maxfd = 0;
		for(i = 0; i < FDSSIZE; i++){
			if (fd[i].fd != -1){
				FD_SET(fd[i].fd, &fds);
				maxfd = ((fd[i].fd > maxfd)? fd[i].fd : maxfd);
			}
		}
		memcpy(&fds_e, &fds, sizeof(fd_set));
		memset(&zero, 0x00, sizeof(struct timeval));
		if ((j = select(maxfd + 1, &fds, NULL, &fds_e, &zero)) == -1){
			for(i = 0; i < FDSSIZE; i++){
				if ((fd[i].fd != -1) && (FD_ISSET(fd[i].fd, &fds_e)))
					fprintf(stderr, "Exception:%d %lu\n", fd[i].fd, t);
			}
			fatal_error("Error in Selecting");
		}
		else
			j = (FD_ISSET(fd[0].fd, &fds)? j - 1 : j);

	//init disconnect fd
	//recv and structure
		t = time(NULL);
		for(i = 1; j > 0; i++){
			if ((fd[i].fd != -1) && (FD_ISSET(fd[i].fd, &fds))){
				j--;
				switch (io[nextrb].flg_i){
				//ringbuf full-buffered
				case 1:
					if (fd[i].io == 0x00){
						switch (recv_chk(fd[i].fd)){
						case 0:
							fprintf(stderr, "disconnected(f):%d %lu\n", fd[i].fd, t);
							break;
						case -1:
							fprintf(stderr, "recv error(f):%d %lu\n", fd[i].fd, t);							
							break;
						default:
							continue;
						}

						close(fd[i].fd);
						init_s_fd(fd + i);
						usedfds--;
						continue;
					}
				//ringbuf NOT full-buffered
				case 0:
					if (fd[i].io == 0x00)
						fd[i].io = io + nextrb;

					switch (recv_req(fd[i].fd, fd[i].io->req, &(fd[i].io->rcvd))){
					case 0:
						fprintf(stderr, "disconnected(e):%d %lu\n", fd[i].fd, t);
						init_s_io(fd[i].io);
						close(fd[i].fd);
						init_s_fd(fd + i);
						usedfds--;
						continue;
					case -1:
						fprintf(stderr, "recv error(e):%d %lu\n", fd[i].fd, t);
						if (fd[i].io == io + nextrb)
							nextrb = (nextrb + 1) % RNGSIZE;
						continue;
					default:
						fprintf(stderr, "recv:%d %lu\n", fd[i].fd, t);
						if (fd[i].io == io + nextrb)
							nextrb = (nextrb + 1) % RNGSIZE;
						break;
					}
				}
				//check crlf2
				//check content-length
				if ((fd[i].io->hbdy = crlf2(fd[i].io->req)) == 0)
					continue;
				else if (content_length(fd[i].io->req, fd[i].io->hbdy, fd[i].io->rcvd) != -1){
				//flg on
					fprintf(stderr, "flg on:%d %lu\n", fd[i].fd, t);
					fd[i].io->fd = fd[i].fd;
					memcpy(&(fd[i].io->addr), &(fd[i].addr), sin6_size);
					fd[i].io->flg_i = 1;
					}
				//non-terminated fd full buffered
				else if (fd[i].io->rcvd == LEN_REQ - 1){
					fprintf(stderr, "full buffered:%d %lu\n", fd[i].fd, t);
					close(fd[i].fd);
					init_s_io(fd[i].io);
				}
				init_s_fd(fd + i);
				usedfds--;
			}
		}

	//accept
		t = time(NULL);
		if (FD_ISSET(fd[0].fd, &fds)){
			i = 0;
			while(usedfds < FDSSIZE){
				c_sock = accept(fd[0].fd, (struct sockaddr *)&c_addr, (socklen_t *)&sin6_size);
				if (c_sock == -1)
					break;
				for(i = i + 1; i < FDSSIZE; i++){
					if (fd[i].fd < 0)
						break;
				}
				fd[i].fd = c_sock;
				fd[i].t_acpt = time(NULL);
				memcpy(&(fd[i].addr), &c_addr, sin6_size);
				usedfds++;
				fprintf(stderr, "accept:%d %lu\n", fd[i].fd, t);
				if (ioctl(fd[i].fd, FIONBIO, &yes) == -1)
					fprintf(stderr, "nbio error:%d %lu\n", fd[i].fd, t);
			}
		}

	//http_io
		t = time(NULL);
		i = nextrb;
		do{
			i = (i - 1 + RNGSIZE) % RNGSIZE;

			if (io[i].flg_i == 0)
				break;

			fprintf(stderr, "process:%d %lu\n", io[i].fd, t);
			nextrec += io_http(io + i, rec + nextrec);
			io[i].flg_o = 1;

			if (nextrec >= RECORD){
				fprintf(stderr, "\nNo memory for new records...\n");
				gshutdown = 0;
				break;
			}
		}while(i != nextrb);

	//send
		t = time(NULL);
		i = nextrb;
		do{
			if (io[i].flg_o == 1){
				fprintf(stderr, "sent:%d %lu\n", io[i].fd, t);
				io[i].sent += send_res(io[i].fd, strlen(io[i].res) - io[i].sent, io[i].res + io[i].sent);
				if (io[i].sent == strlen(io[i].res)){
					close(io[i].fd);
					fprintf(stderr, "close:%d %lu\n", io[i].fd, t);
					init_s_io(io + i);
				}
				else if (io[i].t_res == 0)
					io[i].t_res = t;
				else if (t - io[i].t_res >= TIMEOUT){
					close(io[i].fd);
					fprintf(stderr, "disconnected(s):%d %lu\n", io[i].fd, t);
					init_s_io(io + i);
				}
			}
			i++;
			i %= RNGSIZE;
		}while(i != nextrb);

	//record
		if (usedfds == 1 && save_rec(rec, nextrec))
			fprintf(stderr, "saving...\n");
	}

	if (save_rec(rec, nextrec))
		fprintf(stderr, "saving...\n");

	t = time(NULL);
	close(p_sock);
	fprintf(stderr, "close:%d %lu\n", p_sock, t);
	for(i = 1; i < FDSSIZE; i++){ 
		if (fd[i].fd != -1){
			close(fd[i].fd);
			fprintf(stderr, "close:%d %lu\n", fd[i].fd, t);
		}
	}

	return 0;
}

int recv_chk(int sockfd){
	char buf;

	return (recv(sockfd, &buf, 1, MSG_PEEK));
}

int recv_req(int sockfd, char *buffer, int *rcvd){
	char *ptr;
	int ret, bufsize;

	ptr = buffer + *rcvd;
	bufsize = (LEN_REQ - 1) - *rcvd;

	ret = recv(sockfd, ptr, bufsize, 0);

	if (ret > 0){
		*rcvd += ret;
		ptr += ret;
		*ptr = 0x00;
	}

	return ret;
}

int crlf2(char *buffer){
#define TRM "\r\n\r\n"

	char *ptr;

	ptr = strstr(buffer, TRM);

	if (ptr == NULL)
		return 0;
	else {
		ptr += strlen(TRM);
		return (int)(ptr - buffer);
	}
}

int content_length(char *buffer, int hbdy, int rcvd){
#define CHK "Content-Length:"

	char *ptr;
	int ret;

	ptr = strstr(buffer, CHK);
	if (ptr == NULL)
		ret = 0;
	else{
		ret = (int)strtol(ptr + strlen(CHK), NULL, 10);

		if ((hbdy + ret) > (LEN_REQ - 1))
			ret = (LEN_REQ - 1) - hbdy;

		if (ret != rcvd - hbdy)
			ret = -1;
	}

	return ret;
}

int send_res(int sockfd, int len, char *buffer){
	int ret, bytes_to_send, sent_bytes = 0;

	bytes_to_send = len;
	while(bytes_to_send > 0) {
		ret = send(sockfd, buffer, bytes_to_send, 0);
		if(ret == -1)
			break;
		bytes_to_send -= ret;
		buffer += ret;
		sent_bytes += ret;
	}
	return sent_bytes;
}

void init_s_io(struct io *io){
	io->fd = 0;
	io->flg_i = 0;
	io->flg_o = 0;
	io->rcvd = 0;
	io->hbdy = 0;
	io->sent = 0;
	io->t_res = 0;
	io->req[0] = 0x00;
	io->res[0] = 0x00;
	memset(&(io->addr), 0x00, sizeof(struct sockaddr_in6));
}

void init_s_fd(struct fd *fd){
	memset(fd, 0x00, sizeof(struct fd));
	fd->fd = -1;
}

int io_http(struct io *io, struct record *rec){
	enum method mthd;
	char *req, *res, date[LEN_RFC1123_TIME];
	int i = 0;

	time_t t = time(NULL);
	Rfc1123_DateTime(date, LEN_RFC1123_TIME, &t);
	req = io->req;
	res = io->res;

	//content requested
	if (strncmp(req, "HEAD ", strlen("HEAD ")) == 0){
		if (strncmp(req + strlen("HEAD "), "/list HTTP/", strlen("/list HTTP/")) == 0)
			mthd = HCSV;
		else if (strncmp(req + strlen("HEAD "), "/ HTTP/", strlen("/ HTTP/")) == 0)
			mthd = HEAD;
		else
			mthd = N404;
	}
	else if (strncmp(req, "GET ", strlen("GET ")) == 0){
		if (strncmp(req + strlen("GET "), "/list HTTP/", strlen("/list HTTP/")) == 0)
			mthd = GCSV;
		else if (strncmp(req + strlen("GET "), "/ HTTP/", strlen("/ HTTP/")) == 0)
			mthd = GET;
		else
			mthd = N404;
	}
	else if (strncmp(req, "POST / HTTP/", strlen("POST / HTTP/")) == 0)
		mthd = POST;
	else
		mthd = N404;

	//create header
	if (mthd == N404)
		i = snprintf(res, LEN_RES - i, "%s\r\n", h404);
	else
		i = snprintf(res, LEN_RES - i, "%s\r\n", h200);

	i += snprintf(res + i, LEN_RES - i, "%s\r\n", hsvr);
	i += snprintf(res + i, LEN_RES - i, "Date: %s\r\n", date);

	if (mthd == GCSV || mthd == HCSV || mthd == N404) 
		i += snprintf(res + i, LEN_RES - i, "%s\r\n", hctplain);

	i += snprintf(res + i, LEN_RES - i, "\r\n");

	//create body
	if (mthd == POST){
		//post:req->rec
		proc_post(rec->post, req + io->hbdy);
		rec->t_rec = t;
		memcpy(rec->addr.s6_addr, io->addr.sin6_addr.s6_addr, sizeof(struct in6_addr));
		//print html and post
		i += snprintf(res + i, LEN_RES - i, html, t, rec->post);
		//rec count up
		return 1;
	}
	else if (mthd == GET)
		i += snprintf(res + i, LEN_RES - i, html, t, "");
	else if (mthd == GCSV)
		i += proc_csv(res + i, LEN_RES - i, rec);
	else if (mthd == N404)
		i += snprintf(res + i, LEN_RES - i, "404 Not Found");

	//rec no count
	return 0;
}

void proc_post(char *dst, char *src){
	char *ptrs = src, *ptrd = dst;
	char work[3];
	int len = strlen(src);

	while(*ptrs != 0x00 || ptrd - dst < LEN_PST){
		if ((*ptrs == '%') && (ptrs - src <= len - 3)){
			snprintf(work, 3, "%s", ptrs + 1);
			*ptrd = (char)strtoul(work, NULL, 16);
			ptrs += 2;
		} 
		else 
			*ptrd = *ptrs;

		if ((*ptrd != '\r') && (*ptrd != '\n') && (*ptrd != '\"') && (*ptrd != ','))
			ptrd++;

		ptrs++;
	}
	*ptrd = 0x00;
}

int proc_csv(char *buffer, size_t n, struct record *rec){
	int i = 0;
	char buf[INET6_ADDRSTRLEN];

	if (rec->flg != 0)
		i += snprintf(buffer + i, n - i, "NO DATA");
	else {
		while(rec->flg == 0 && i != n - 1){
			rec--;
			i += snprintf(buffer + i, n - i, "%s,", inet_ntop(AF_INET6, rec->addr.s6_addr, buf, INET6_ADDRSTRLEN));
			i += snprintf(buffer + i, n - i, "%s,", rec->post);
			i += snprintf(buffer + i, n - i, "%s", ctime(&(rec->t_rec)));
		}
	}

   return i;
}

int save_rec(struct record *rec, int next){
	static int saved = -1;
	int i;
	char buf[INET6_ADDRSTRLEN];

	if (saved != next - 1){
		for(i = saved + 1; i < next; i++){
			fprintf(stdout, "%s,", inet_ntop(AF_INET6, rec[i].addr.s6_addr, buf, INET6_ADDRSTRLEN));
			fprintf(stdout, "%s,", rec[i].post);
			fprintf(stdout, "%s", ctime(&(rec[i].t_rec)));
		}

		fflush(stdout);
		saved = i - 1;

		return 1;
	}
	else
		return 0;
}

int Rfc1123_DateTime(char *buffer, size_t n, time_t *t){
    struct tm tm;
    char buf[LEN_RFC1123_TIME];

    gmtime_r(t, &tm);

    strftime(buf, LEN_RFC1123_TIME, "---, %d --- %Y %H:%M:%S GMT", &tm);
    memcpy(buf, DAY_NAMES[tm.tm_wday], 3);
    memcpy(buf+8, MONTH_NAMES[tm.tm_mon], 3);

    return (snprintf(buffer, n, "%s", buf));
}

void fatal_error(const char *message){
	char errmsg[100];

	memset(errmsg, 0x00, 100);
	strcpy(errmsg, "[!!] Fatal Error ");
	strncat(errmsg, message, 82);
	perror(errmsg);
	exit(-1);
}