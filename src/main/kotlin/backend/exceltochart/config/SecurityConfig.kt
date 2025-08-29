package backend.exceltochart.config
import org.springframework.context.annotation.Bean
import org.springframework.context.annotation.Configuration
import org.springframework.security.config.annotation.web.builders.HttpSecurity
import org.springframework.security.web.SecurityFilterChain

@Configuration
class SecurityConfig(
    private val corsConfigurationSource: org.springframework.web.cors.CorsConfigurationSource
) {

    @Bean
    fun filterChain(http: HttpSecurity): SecurityFilterChain {
        http
            .csrf { it.disable() } // CSRF 비활성화 (API 서버면 보통 꺼둠)
            .cors { it.configurationSource(corsConfigurationSource) } // CORS 적용
            .authorizeHttpRequests { auth ->
                auth
                    .requestMatchers("/**").permitAll() // 모든 요청 허용
                //.anyRequest().authenticated()      // → 나중에 인증 붙일 때 이걸로 변경
            }

        return http.build()
    }
}
