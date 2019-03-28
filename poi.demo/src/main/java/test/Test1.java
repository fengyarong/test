package test;

import java.util.Date;

import redis.clients.jedis.Jedis;

public class Test1 {
	public static void main(String[] args) {
		Jedis j =new Jedis("localhost",6379);
		j.connect();
		j.set("zhangsan","22");
		j.disconnect();
	}
}
